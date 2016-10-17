
# <a name="error.message-property"></a>Error.message 属性
获取错误的详细描述。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含 Selection 最后一次更改的版本**|1.1|

```
var errMessage = asyncResult.error.message;
```


## <a name="return-value"></a>返回值

以 **字符串** 形式的错误描述。


## <a name="remarks"></a>备注

**Error** 对象及其属性可从 [AsyncResult](../../reference/shared/asyncresult.md) 对象进行访问，后者在作为异步数据操作的 _callback_ 自变量传递的函数中返回。


## <a name="example"></a>示例

若要导致引发错误，选择表或矩阵，然后调用  `setText` 函数。


```js
function setText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            if (asyncResult.status === "failed")
                var error = asyncResult.error;
            write(error.name + ": " + error.message);
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

||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
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
|1.1|增加了对 Access 相关内容外接程序的支持。|
|1.0|引入|
