
# <a name="asyncresult.asynccontext-property"></a>AsyncResult.asyncContext 属性
获取与传入时状态相同的传递给调用方法的可选 _asyncContext_ 参数的用户定义项。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```
var myContext = asynchResult.asyncContext;
```


## <a name="return-value"></a>返回值

返回传递给调用方法的可选  **asyncContext** 参数的用户定义项（该项可以是任何 JavaScript 类型： **String**、 **Number**、 **Boolean**、 **Object**、 **Array**、 **Null** 或 _Undefined_）。如果您未向  **asyncContext** 参数传递任何东西，则返回 _Undefined_ 。


## <a name="example"></a>示例




```js
function getDataWithContext() {
    var format = "Your data: ";
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { asyncContext: format }, showDataWithContext);
}

 function showDataWithContext(asyncResult) {
    write(asyncResult.asyncContext + asyncResult.value);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**适用于 Mac 的 Outlook**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**项目**||||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
