
# <a name="projectdocument.settaskfieldasync-method-(javascript-api-for-office-v1.1)"></a>ProjectDocument.setTaskFieldAsync 方法（适用于 Office 的 JavaScript API v1.1）
异步设置指定任务的指定字段的值。 **重要说明：**此 API 仅可在 Windows 桌面上的 Project 2016 中运行。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**添加内容的版本**|1.1|

```js
Office.context.document.setTaskFieldAsync(taskId, fieldId, fieldValue[, options][, callback]);
```


## <a name="parameters"></a>参数


_taskId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;任务的 GUID。必需。<br/><br/>
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目标字段的 ID，作为 [ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)常量或其对应的整数值。必需。<br/><br/>
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目标字段的值为，作为**字符串**、**数字**、**布尔值**或**对象**。必需。<br/><br/>
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;以下是 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：<br/><br/>

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型：**数组、布尔值、null、数字、对象、字符串**或**未定义**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[AsyncResult](../../reference/shared/asyncresult.md) 对象中未经改动的返回的任何类型的用户定义项。可选。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，可以使用 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}` 格式传递 _asyncContext_ 参数。<br/><br/>
_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**函数**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;返回方法调用时调用的函数，其唯一的参数的类型为 [AsyncResult](../../reference/shared/asyncresult.md)。可选。
    

## <a name="callback-value"></a>回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **setTaskFieldAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。



|**名称**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选 _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[error](../../reference/shared/asyncresult.error.md)|关于错误的信息（如果 **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的 **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|此方法不返回值。|

## <a name="remarks"></a>备注

首先调用 [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 或 [getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md) 方法以获取任务 GUID，然后将 GUID 作为 _taskId_ 参数传递到 **setTaskFieldAsync**。每次异步调用中仅可更新一个任务的一个字段。


## <a name="example"></a>示例

以下代码示例调用 [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 以获取在任务视图中当前所选任务的 GUID。然后它将通过递归调用 **setTaskFieldAsync** 设置两个任务字段值。

示例中使用的 [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 方法要求任务视图（例如，任务与使用情况）为活动视图，且任务已选中。有关根据活动视图类型激活按钮的示例，请参阅 [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) 方法。

示例假定您的外接程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义。




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function setTaskInfo() {
        getTaskGuid().then(
            function (data) {
                setTaskFields(data);
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

    // Set the specified fields for the selected task.
    function setTaskFields(taskGuid) {
        var targetFields = [Office.ProjectTaskFields.Active, Office.ProjectTaskFields.Notes];
        var fieldValues = [true, 'Notes for the task.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setTaskFieldAsync(
                taskGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**||
|**最低权限级别**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|引入|

## <a name="see-also"></a>另请参阅



#### <a name="other-resources"></a>其他资源


[getSelectedTaskAsync 方法](../../reference/shared/projectdocument.getselectedresourceasync.md)
[getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md)
[AsyncResult 对象](../../reference/shared/asyncresult.md)
[ProjectTaskFields 枚举](../../reference/shared/projecttaskfields-enumeration.md)
[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
