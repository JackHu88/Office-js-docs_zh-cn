
# <a name="tablebinding.hasheaders-property"></a>TableBinding.hasHeaders 属性
获取表是否具有标题。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint、Project、Word|
|在**要求集[中可用](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**包含 Selection 最后一次更改的版本**|1.1|

```
var colCount = bindingObj.hasHeaders;
```


## <a name="return-value"></a>返回值

如果指定的 [TableBinding](../../reference/shared/binding.tablebinding.md) 具有标题，则返回 **true**；否则返回 **false**。


## <a name="example"></a>示例




```js
function showBindingHasHeaders() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Binding has headers: " + asyncResult.value.hasHeaders);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录





****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|添加了对 Access 相关应用程序中此事件的支持。|
|1.0|引入|
