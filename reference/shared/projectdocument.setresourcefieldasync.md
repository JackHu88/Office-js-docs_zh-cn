

# ProjectDocument.setResourceFieldAsync 方法
异步设置指定资源的指定字段的值。
 **重要说明：**此 API 仅可在 Windows 桌面上的 Project 2016 中运行。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**在其中添加**|1.1|

```js
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## 参数

_resourceId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;资源的 GUID。 必需。
    
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目标字段的 ID，作为 [ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md) 常量或其对应的整数值。 必需。
    
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目标字段的值为，作为**字符串**、**数字**、**布尔值**或**对象**。 必需。
    
_选项_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;以下是 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型：**数组、布尔值、null、数字、对象、字符串**或**未定义**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[AsyncResult](../../reference/shared/asyncresult.md) 对象中未经改动的返回的任何类型的用户定义项。 可选。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，可以使用 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}` 格式传递 _asyncContext_ 参数。


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**函数**

&nbsp;&nbsp;&nbsp;&nbsp;返回方法调用时调用的函数，其唯一的参数的类型为 [AsyncResult](../../reference/shared/asyncresult.md)。 可选。

    

## 回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **setResourceFieldAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


|**名称**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选  _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[错误](../../reference/shared/asyncresult.error.md)|关于错误的信息（ 如果  **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的  **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|此方法不返回值。|

## 备注

首先调用 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 或 [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) 方法以获取资源 GUID，然后将 GUID 作为 _resourceId_ 参数传递到 **setResourceFieldAsync**。每次异步调用中仅可更新一个资源的一个字段。


## 示例

以下代码示例调用 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 以获取在资源视图中当前所选资源的 GUID。然后它将通过递归调用 **setResourceFieldAsync** 设置两个资源字段值。

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
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
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

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
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


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**||
|**最低权限级别**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录

|**版本**|**更改内容**|
|:-----|:-----|
|1.1|引入|

## 另请参阅



#### 其他资源


[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
(#getselectedresourceasync)[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)
(#getresourcebyindexasync)[AsyncResult 对象](../../reference/shared/asyncresult.md)
(#asyncresult-对象)[ProjectResourceFields 枚举](../../reference/shared/projectresourcefields-enumeration.md)
(#projectresourcefields-枚举)[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)

