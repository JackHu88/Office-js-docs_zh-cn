
# ProjectDocument.getProjectFieldAsync 方法
异步获取活动项目中指定字段的值。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**在其中添加**|1.0|

```
Office.context.document.getProjectFieldAsync(fieldId[, options][, callback]);
```


## 参数



|**名称**|**类型**|**说明**|
|:-----|:-----|:-----|:-----|
| _fieldId_|[ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)|目标字段的 ID。必需。|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。|
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。|
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。|

## 回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getProjectFieldAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


****


|**Name**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选  _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[错误](../../reference/shared/asyncresult.error.md)|关于错误的信息（ 如果  **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的  **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|包含  **fieldValue** 属性，它表示指定字段的值。|

## 示例

以下代码示例获取活动项目的三个指定字段的值，然后将值显示在应用程序中。

前一个调用成功返回后，示例将递归调用 **getProjectFieldAsync**。它还会跟踪调用以确定所有调用何时发送完。

示例假定您的应用程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义以下页面控件。




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information for the active project.
            getProjectInformation();
        });
    };

    // Get the specified fields for the active project.
    function getProjectInformation() {
        var fields =
            [Office.ProjectProjectFields.Start, Office.ProjectProjectFields.Finish, Office.ProjectProjectFields.GUID];
        var fieldValues = ['Start: ', 'Finish: ', 'GUID: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == fields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }
            else {
                Office.context.document.getProjectFieldAsync(
                    fields[index],
                    function (result) {

                        // If the call is successful, get the field value and then get the next field.
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
|**在要求集中可用**||
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


[ProjectProjectFields 枚举](../../reference/shared/projectprojectfields-enumeration.md)

[AsyncResult 对象](../../reference/shared/asyncresult.md)

[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
