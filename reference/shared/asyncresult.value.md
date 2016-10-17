
# <a name="asyncresult.value-property"></a>AsyncResult.value 属性
获取此异步操作的负载或内容（如有）。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```js
var dataValue = asyncResult.value;
```


## <a name="return-value"></a>返回值

返回做出异步调用时请求的值。 


 >**注意**：**value** 属性为特定 "Async" 方法返回的内容会有所不同，具体取决于该方法的目的和上下文。若要确定 **value** 属性为 "Async" 方法返回了什么内容，请参阅该方法主题的“回调值”部分。有关 "Async" 方法的完整列表，请参阅 [AsyncResult](../../reference/shared/asyncresult.md) 对象主题的“备注”部分。


## <a name="remarks"></a>注解

访问函数中作为实参传递给 "Async" 方法的 **callback** 形参的 _AsyncResult_ 对象，如 [Document](../../reference/shared/document.getselecteddataasync.md) 对象的 [getSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) 和 **setSelectedDataAsync** 方法。


## <a name="example"></a>示例




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**适用于 Mac 的 Office**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**项目**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
