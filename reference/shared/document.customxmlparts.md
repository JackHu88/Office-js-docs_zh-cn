
# <a name="document.customxmlparts-property"></a>Document.customXmlParts 属性
获取文档中表示自定义 XML 部件的对象。

|||
|:-----|:-----|
|**主机：**|Word|
|**添加内容的版本**|1.1|

```js
var xmlParts = Office.context.document.customXmlParts;
```


## <a name="return-value"></a>返回值

一个 [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) 对象。


## <a name="example"></a>示例




```js
function getCustomXmlParts(){
    Office.context.document.customXmlParts.getByNamespaceAsync('http://tempuri.org', function (asyncResult) {
        write('Retrieved ' + asyncResult.value.length + ' custom XML parts');
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 上 Word 的支持|
|1.0|引入|
