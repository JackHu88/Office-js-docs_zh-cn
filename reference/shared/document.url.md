
# <a name="document.url-property"></a>Document.url 属性
获取主机应用程序当前打开的文档的 URL。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Project 和 Word|
|**包含最后一次更改的版本**|1.1|

```
var docUrl = Office.context.document.url;
```


## <a name="return-value"></a>返回值

文档的 URL。如果 URL 不可用，则返回  **null**。


## <a name="remarks"></a>注解

 **重要信息：** **url** 属性返回信息，其中可能会在文档名称和存储位置中包含个人身份信息 (PII)。如果必须存储或传输这些信息，请务必以加密格式执行此操作。


## <a name="example"></a>示例




```
function displayDocumentUrl() {
    write(Office.context.document.url);
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
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录





****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Word Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关内容外接程序的支持。|
|1.0|引入|
