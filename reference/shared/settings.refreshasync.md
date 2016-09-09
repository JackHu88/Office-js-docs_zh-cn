

# Settings.refreshAsync 方法
读取文档中保存的所有设置并刷新内容或任务窗格外接程序在内存中保留的这些设置的副本。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|Settings|
|**包含最后一次更改的版本**|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## 参数

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**对象**

&nbsp;&nbsp;&nbsp;&nbsp;返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult**。

    



## 回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给  **refreshAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问包含刷新后的值的 [Settings](../../reference/shared/settings.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

此方法适用于采用共同创作方案的 Word 和 PowerPoint，即相同外接程序的多个实例在处理同一个文档。由于各个外接程序处理的是在用户打开文档时从文档中加载的设置的内存中副本，因此每个用户使用的设置值可能会不同步。只要外接程序实例调用 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法，将用户的所有设置都保留到文档中，就会出现这种情况。从外接程序的 **settingsChanged** 事件处理程序调用 [refreshAsync](../../reference/shared/settings.settingschangedevent.md) 方法会刷新所有用户的设置值。

可从为 Excel 创建的 外接程序 调用  **refreshAsync**方法，但是因为它们不支持合著，因此没有理由那么做。


## 示例




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
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
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|Settings|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关内容外接程序中自定义设置的支持。|
|1.0|引入|
