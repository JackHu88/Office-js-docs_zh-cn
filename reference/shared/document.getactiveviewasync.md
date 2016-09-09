
# Document.getActiveViewAsync 方法
 返回演示文稿（编辑或读取）的当前视图的状态。

|||
|:-----|:-----|
|**主机：**Excel、PowerPoint 和 Word|**外接程序类型：**内容、任务窗格|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|ActiveView|
|**在 ActiveView 中添加**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **getActiveViewAsync** 方法的回调函数中，[AsyncResult.value](../../reference/shared/asyncresult.value.md) 属性返回演示文稿的当前视图的状态。 返回的值可以是 `edit` 或 `read`。  `edit` 对应于任何你可以从中编辑幻灯片的视图，比如“**常规**”或“**大纲视图**”。  `read` 对应于“**幻灯片放映**”或“**阅读视图**”。


## 注解

当视图更改时可以触发事件。


## 示例

若要获得当前演示文稿的视图，您需要编写一个可返回该值的回调函数。以下示例显示如何：


-  **将可返回视图类型的匿名回调函数传递给** _getActiveViewAsync_ 方法的 **callback** 参数。
    
-  在外接程序页面上 **显示该值**。
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|||Y|

|||
|:-----|:-----|
|**在要求集中可用**|ActiveView|
|**在 ActiveView 中添加**|1.1|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录





****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|引入。|
