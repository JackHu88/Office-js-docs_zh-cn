
# <a name="customxmlnode.settextasync-method"></a>CustomXmlNode.setTextAsync 方法
异步设置自定义 XML 部件中 XML 节点的文本。

|||
|:-----|:-----|
|**主机：**|Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|CustomXmlParts|
|**添加内容的版本**|1.2|

```
customXmlNodeObj.setTextAsync(text, [asyncContext,]callback(asyncResult);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|
|:-----|:-----|:-----|
| _text_|**string**|必需。XML 节点的文本值。|
| _asyncContext_|**object**|可选。用户定义的对象，适用于 [AsyncResult](../../reference/shared/asyncresult.md) 对象的 asyncContext 属性。当回调为命名的函数时，使用此选项为 **AsyncResult** 提供对象或值。|
| _callback_|**object**|可选。返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。|

## <a name="callback-value"></a>回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **setTextAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|未使用。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|指示操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。如果尚未设置 _asyncContext_，则此属性将返回其未定义的形式。|

## <a name="example"></a>示例

了解如何在自定义 XML 部件中设置节点的文本值。


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to the host.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:subject", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. This will be the subject element in this example.
                var subjectNode = getNodesAsyncResult.value[0];
                
                // Set the text value of the subject node and use the asyncContext. This results in a call to the host. 
                // The results are logged to the browser console. 
                subjectNode.setTextAsync("newSubject", {asyncContext: "StateNormal"}, function (setTextAsyncResult) {
                   console.log("The status of the call: " + setTextAsyncResult.status);
                   console.log("The asyncContext value = " + setTextAsyncResult.asyncContext);
                });
            });
        });
    });
});
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
|1.1|添加 setTextAsync。|
