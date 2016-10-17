
# <a name="office.initialize-event"></a>Office.initialize 事件
加载运行时环境和外接程序准备好开始与应用和托管文档交互时发生。 

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## <a name="remarks"></a>备注

_initialize_ 事件侦听器函数的 **reason** 参数返回 [InitializationReason](../../reference/shared/initializationreason-enumeration.md) 枚举值，以指定初始化的发生方式。任务窗格外接程序或内容外接程序可以通过下列两种方式进行初始化：


- 用户只需从 Office 主机应用程序中功能区的“**插入**”选项卡上的“**外接程序**”下拉列表的“**最近使用的外接程序**”部分或从“**插入外接程序**”对话框插入该任务窗格或内容外接程序。
    
- 用户打开已包含外接程序的文档。
    

 >**注意**：**initialize** 事件侦听器函数的 reason 参数仅为任务窗格外接程序和内容外接程序返回 **InitializationReason** 枚举值，不会为 Outlook 外接程序返回值。


## <a name="example"></a>示例

您可以使用  **InitializationEnumeration** 的值为第一次插入外接程序与外接程序已是文档的一部分这两种情况实现不同的逻辑。以下示例显示使用 _reason_ 参数的值以显示任务窗格或内容外接程序的初始化方式的一些简单逻辑。


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此事件。空的单元格表示相应的 Office 主机应用程序不支持此事件。

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
|**外接程序类型**|内容、Outlook、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对初始化 Access 相关内容外接程序的支持。|
|1.0|引入|
