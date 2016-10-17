
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>在 Outlook 中撰写约会时获取或设置位置

适用于 Office 的 JavaScript API 提供了异步方法（[getAsync](../../reference/outlook/Location.md) 和 [setAsync](../../reference/outlook/Location.md)）以获取和设置用户正在撰写的约会的位置。这些异步方法仅对撰写外接程序可用。若要使用这些方法，请确保已将 Outlook 的外接程序清单相应地设置为在撰写窗体中激活外接程序，如[创建适用于撰写窗体的 Outlook 外接程序](../outlook/compose-scenario.md)所述。

[location](../../reference/outlook/Office.context.mailbox.item.md) 属性适用于约会撰写和阅读窗体中的读取权限。在阅读窗体中，您可以从父对象直接访问此属性，如：




```js
item.location
```

但在撰写窗体中，由于用户和外接程序可同时插入或更改位置，您必须使用异步方法  **getAsync** 获取位置，如下所示：




```js
item.location.getAsync
```

**location** 属性仅可用于约会的撰写窗体中（而不能用于阅读窗体中）的写入权限。

与适用于 Office 的 JavaScript API 中的大多数异步方法相同，**getAsync** 和 **setAsync** 采用可选输入参数。有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md)。


## <a name="to-get-the-location"></a>获取位置


本节演示获取用户正在撰写的约会的位置、并显示位置的代码示例。此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会激活外接程序，如下所述。


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

若要使用  **item.location.getAsync**，可提供一个检查异步调用状态和结果的回调方法。您可以通过  _asyncContext_ 可选形参向回调方法提供任何必要实参。可以使用回调的输出形参 _asyncResult_ 获取状态、结果和任何错误。如果异步调用成功，则可以使用 [AsyncResult.value](../../reference/outlook/simple-types.md) 属性获取字符串形式的位置。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-set-the-location"></a>设置位置


本节演示设置用户正在撰写的约会的位置的代码示例。与上一示例类似，此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会激活外接程序。

若要使用  **item.location.setAsync**，可在数据形参中指定一个最多 255 字符的字符串。也可以在  _asyncContext_ 形参中为回调方法提供一个回调方法和任何实参。您应在回调的 _asyncResult_ 输出形参中检查状态、结果和所有错误消息。如果异步调用成功， **setAsync** 会将指定的位置字符串作为纯文本插入，并覆盖该项目的任何现有位置。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="additional-resources"></a>其他资源



- [在 Outlook 的撰写窗体中获取和设置项目数据](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](../outlook/item-data.md)
    
- [创建适用于撰写窗体的 Outlook 外接程序](../outlook/compose-scenario.md)
    
- [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [在 Outlook 中撰写约会或邮件时获取、设置或添加收件人](../outlook/get-set-or-add-recipients.md)
    
- [在 Outlook 中撰写约会或邮件时获取或设置主题](../outlook/get-or-set-the-subject.md)
    
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](../outlook/insert-data-in-the-body.md)
    
- [在 Outlook 中撰写约会时获取或设置时间](../outlook/get-or-set-the-time-of-an-appointment.md)
    
