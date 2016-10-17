
# <a name="customxmlparts.addasync-method"></a>CustomXmlParts.addAsync 方法
将新的自定义 XML 部件异步添加到文件中。

|||
|:-----|:-----|
|**主机：**|Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|CustomXmlParts|
|**添加内容的版本**|1.1|

```
Office.context.document.customXmlParts.addAsync(xml [, options], callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _xml_|**string**|要添加到新创建的自定义 XML 部件的 XML。必需。||
| _options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|返回新创建的 [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="example"></a>示例




```js
function addXMLPart() {
    Office.context.document.customXmlParts.addAsync('<root categoryId="1" xmlns="http://tempuri.org"><item name="Cheap Item" price="$193.95"/><item name="Expensive Item" price="$931.88"/></root>', function (result) {
        });
}
```








```
function addXMLPartandHandler() {
    Office.context.document.customXmlParts.addAsync("<testns:book xmlns:testns='http://testns.com'><testns:page number='1'>Hello</testns:page><testns:page number='2'>world!</testns:page></testns:book>",
        function(r) { r.value.addHandlerAsync(Office.EventType.DataNodeDeleted,
            function(a) {write(a.type)
            },
                function(s) {write(s.status)
                });
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|CustomXmlParts|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Word 的支持。|
|1.0|引入|
