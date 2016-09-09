
# ProjectDocument.getResourceFieldAsync 方法
异步获取资源视图中指定资源的指定字段的值。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**在其中添加**|1.0|

```
Office.context.document.getResourceFieldAsync(resourceId, fieldId[, options][, callback]);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _resourceId_|**string**|资源的 GUID。必需。||
| _fieldId_|[ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)|目标字段的 ID。必需。||
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getResourceFieldAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


****


|**姓名**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选  _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[错误](../../reference/shared/asyncresult.error.md)|关于错误的信息（ 如果  **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的  **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|包含  **fieldValue** 属性，它表示指定字段的值。|

## 备注

首先调用 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 方法获取资源 GUID，然后将其作为 _resourceId_ 参数传递到 **getResourceFieldAsync**。如果活动视图不是资源视图（例如“资源使用情况”或“资源表”视图），或者如果在资源视图中未选择资源，[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 将返回 5001 错误（内部错误）。请参阅 [addHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md)，查看使用 [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) 事件和 [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) 方法以基于活动视图类型激活按钮的示例。


## 示例

以下代码示例调用 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 以获取在资源视图中当前所选资源的 GUID。然后它将通过递归调用 **getResourceFieldAsync** 获取三个资源字段值。

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
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the GUID of the resource and then get the resource fields.
    function getResourceInfo() {
        getResourceGuid().then(
            function (data) {
                getResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
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

    // Get the specified fields for the selected resource.
    function getResourceFields(resourceGuid) {
        var targetFields =
            [Office.ProjectResourceFields.Name, Office.ProjectResourceFields.Units, Office.ProjectResourceFields.BaseCalendar];
        var fieldValues = ['Name: ', 'Units: ', 'Base calendar: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // If the call is successful, get the field value and then get the next field.
            else {
                Office.context.document.getResourceFieldAsync(
                    resourceGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
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


[getSelectedResourceAsync 方法](../../reference/shared/projectdocument.getselectedresourceasync.md)

[ProjectResourceFields 枚举](../../reference/shared/projectresourcefields-enumeration.md)

[AsyncResult 对象](../../reference/shared/asyncresult.md)

[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
