
# ProjectDocument.getSelectedTaskAsync 方法
在任务视图中异步获取所选任务的 GUID。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**在其中添加**|1.0|

```
Office.context.document.getSelectedTaskAsync([options,] [callback]);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getSelectedTaskAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


****


|**Name**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选  _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[错误](../../reference/shared/asyncresult.error.md)|关于错误的信息（ 如果  **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的  **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|**string** 形式的所选任务的 GUID。|

## 备注

在 Project 外接程序中，任务 GUID 比任务标识号（例如，在甘特图中，第一个任务的标识号为 **1**）更有用。任务 GUID 可用于访问 Project 任务信息，如与可见模式下的 Project Server 同步的 SharePoint 项目中的任务。你还可以将任务 GUID 保存到本地变量中，以供 [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md) 和 [getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md) 方法使用。

如果活动视图不是任务视图（例如，“甘特图”或“任务分配状况”视图），或者未选择任务视图中的任何任务，则 **getSelectedTaskAsync** 返回 5001 错误（内部错误）。有关使用 [ViewSelectionChanged](../../reference/shared/projectdocument.addhandlerasync.md) 事件和 [getSelectedViewAsync](../../reference/shared/projectdocument.viewselectionchanged.event.md) 方法根据活动视图的类型激活按钮的示例，请参阅 [addHandlerAsync 方法](../../reference/shared/projectdocument.getselectedviewasync.md)。


## 示例

以下代码示例调用 **getSelectedTaskAsync**，获取任务视图中当前选定的任务的 GUID。然后，它通过调用 [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md) 获取任务属性。

示例假定您的外接程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getTaskInfo);
        });
    };

    // // Get the GUID of the task, and then get local task properties.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskProperties(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get local properties for the selected task, and then display it in the add-in.
    function getTaskProperties(taskGuid) {
        Office.context.document.getTaskAsync(
            taskGuid,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var taskInfo = result.value;
                    var output = String.format(
                        'Name: {0}<br/>GUID: {1}<br/>SharePoint task ID: {2}<br/>Resource names: {3}',
                        taskInfo.taskName, taskGuid, taskInfo.wssTaskId, taskInfo.resourceNames);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();

    


```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**|Selection|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## 另请参阅



#### 其他资源


[getTaskAsync 方法](../../reference/shared/projectdocument.gettaskasync.md)

[AsyncResult 对象](../../reference/shared/asyncresult.md)

[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
