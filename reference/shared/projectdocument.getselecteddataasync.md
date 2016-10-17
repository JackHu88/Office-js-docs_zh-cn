
# <a name="projectdocument.getselecteddataasync-method"></a>ProjectDocument.getSelectedDataAsync 方法
异步获取甘特图视图中一个或多个单元格的当前选择中包含的数据文本值。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**添加内容的版本**|1.0|

```
Office.context.document.getSelectedDataAsync(coercionType[, options][, callback]);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)|要返回的数据结构的类型。必需。<br/>Project 2013 仅支持 **Office CoercionType.Text** 或 `"text"`。||
| _options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|数字或日期值的格式。<br/>Project 2013 会忽略此参数，并在内部将其设置为 `unformatted`。||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|指定是仅包含可见数据，还是包含所有数据。 <br/>Project 2013 会忽略此参数，并在内部将其设置为 `all`。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getSelectedDataAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


****


|**名称**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选  _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[error](../../reference/shared/asyncresult.error.md)|关于错误的信息（如果 **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的 **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|选定单元格的文本值。|

## <a name="remarks"></a>备注

**ProjectDocument.getSelectedDataAsync** 方法替代 [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法，并返回“甘特图”视图中一个或多个单元格内选定数据的文本值。**ProjectDocument.getSelectedDataAsync** 仅支持文本格式的 [CoercionType](../../reference/shared/coerciontype-enumeration.md)，不支持 `matrix`、`table` 或其他格式。


## <a name="example"></a>示例

以下代码示例获取选定单元格的值。它使用可选的  _asyncContext_ 参数将部分文本传递到回调函数。

示例假定您的应用程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义以下页面控件。




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
            $('#get-info').click(getSelectedText);
        });
    };

    // Get the text from the selected cells in the document, and display it in the add-in.
    function getSelectedText() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            {asyncContext: 'Some related info'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'Selected text: {0}<br/>Passed info: {1}',
                        result.value, result.asyncContext);
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


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**|Selection|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## <a name="see-also"></a>另请参阅



#### <a name="other-resources"></a>其他资源


[AsyncResult 对象](../../reference/shared/asyncresult.md)

[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md)

[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
