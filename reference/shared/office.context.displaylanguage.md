
# <a name="context.displaylanguage-property"></a>Context.displayLanguage 属性
获取用户针对 Office 主机应用程序的 UI 指定的区域设置（语言）。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```
var myDisplayLanguage = Office.context.displayLanguage;
```


## <a name="return-value"></a>返回值

RFC 1766 语言标记格式的一个  **string**，如  `en-US`。


## <a name="remarks"></a>注解

The  **displayLanguage** value reflects the current **Display Language** setting specified with **File** > **Options** > **Language** in the Office host application.

在 Access Web 应用程序相关内容外接程序中， **displayLanguage** 属性会获取外接程序语言（例如，"en-US"）。


## <a name="example"></a>示例




```js
function sayHelloWithDisplayLanguage() {
    var myDisplayLanguage = Office.context.displayLanguage;
    switch (myDisplayLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y|||
|**Excel**|Y|Y|Y||
|**Outlook**|Y|Y||Y|
|**PowerPoint**|Y|Y|Y||
|**Project**|Y||||
|**Word**|Y|Y|Y||

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
|1.1|增加了对 Access 相关内容外接程序中此 API 的访问权限。|
|1.0|引入|
