
# <a name="document.gotobyidasync-method"></a>Document.goToByIdAsync 方法
转到文档中指定的对象或位置。

|||
|:-----|:-----|
|**主机：**|Excel、PowerPoint 和 Word|
|**在要求集中可用**|不在集合中|
|**添加内容的版本**|1.1|

```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _id_|**string** 或 **number**|要转到的对象或位置的标识符。必需。||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|要转到的位置类型。必需。||
| _options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|指定是否选择（突出显示）由 _id_ 参数指定的位置。|**在 Excel 中：**<br/> **Office.SelectionMode.Selected** 选择绑定中或已命名项目中的所有内容。 <br/>**Office.SelectionMode.None** 对于文本绑定，选择单元格；对于矩阵绑定、表绑定和已命名的项目，选择第一个数据单元格（不是表格标题行中的第一个单元格）。<br/><br/> **在 PowerPoint 中：**<br/> **Office.SelectionMode.Selected** 选择幻灯片标题或幻灯片上的第一个文本框。<br/> **Office.SelectionMode.None** 不选择任何内容。<br/><br/> **在 Word 中：**<br/> **Office.SelectionMode.Selected** 选择绑定中的所有内容。 <br/>**Office.SelectionMode.None** 对于文本绑定，将光标移到文本开头；对于矩阵绑定和表绑定，选择第一个数据单元格（不是表格标题行中的第一个单元格）。|
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **goToByIdAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|返回当前视图。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>备注

PowerPoint 不支持“**母版视图**”中的 **goToByIdAsync** 方法。


## <a name="example"></a>示例

 **按 ID 转到绑定（Word 和 Excel）**

以下示例显示如何：


-  使用 **addFromSelectionAsync** 方法[创建表绑定](../../reference/shared/bindings.addfromselectionasync.md)，作为要使用的示例绑定。
    
-  **指定该绑定**作为要转至的绑定。
    
-  **将返回操作状态的匿名回调函数传递给** _goToByIdAsync_ 方法的 **callback** 参数。
    
-  在外接程序页面上 **显示该值**。
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **转到电子表格中的表 (Excel)**

以下示例显示如何：


-  **按名称指定一个表**作为要转到的表。
    
-  **将返回操作状态的匿名回调函数传递给** _goToByIdAsync_ 方法的 **callback** 参数。
    
-  在外接程序页面上 **显示该值**。
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **按 ID 转到当前选定的幻灯片 (PowerPoint)**

以下示例显示如何：


-  使用 **getSelectedDataAsync** 方法[获取当前所选的幻灯片的 ID](../../reference/shared/document.getselecteddataasync.md)。
    
-  **指定返回的 id** 作为要转到的幻灯片。
    
-  **将返回操作状态的匿名回调函数传递给** _goToByIdAsync_ 方法的 **callback** 参数。
    
-  **在外接程序的页面上显示**`asyncResult.value` 返回的字符串化 JSON 对象的值，其中包含有关所选幻灯片的信息。
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **按索引转到幻灯片 (PowerPoint)**

以下示例显示如何：


-  **指定要转到的第一张、最后一张、上一张或下一张幻灯片的索引**。
    
-  **将返回操作状态的匿名回调函数传递给** _goToByIdAsync_ 方法的 **callback** 参数。
    
-  在外接程序页面上 **显示该值**。
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

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
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|引入|
