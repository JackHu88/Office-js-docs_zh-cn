
# 在 Outlook 的撰写窗体中添加和删除项目附件

您可以使用 [addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 和 [addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法分别将文件和 Outlook 项目附加到用户撰写的项目。两种方法均为异步方法，这意味着执行可以继续，而无需等待 add-attachment 操作完成。根据所添加附件的原始位置和大小，add-attachment 异步调用可能需要一会才能完成。如果有些任务依赖于要完成的操作，您应该在回调方法中执行这些任务。此回调方法为可选，将在附件上载完成时调用。回调方法将 [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.md) 对象作为输出参数，提供任何状态、错误以及从 add-attachment 操作返回的值。如果回调需要任何额外的参数，您可以在可选的 _options.aysncContext_ 参数中进行指定。 _options.asyncContext_ 可以为您的回调方法希望的任何类型。

例如，你可以将 _options.asyncContext_ 定义为一个 JSON 对象，该对象包含一个或多个键值对，其中用“:”字符分隔键和值，用“,”分隔键值对。 你可以找到有关 [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md) 中的 Office 外接程序平台中的 [将可选参数传递给异步方法](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) 的更多示例。 下面的示例演示了如何使用 **asyncContext** 参数将 2 个自变量传递给回调方法：




```js
{ asyncContext: { var1: 1, var2: 2} }
```

您可以使用  **AsyncResult** 对象的 **status** 和 **error** 属性，检查回调方法中异步方法调用是成功还是出现错误。如果附加成功完成，您可以使用 **AsyncResult.value** 属性获取附件 ID。附件 ID 是一个证书，您稍后可使用附件 ID 删除附件。


 >**注释**  作为最佳做法，您应仅在同一外接程序在同一会话中添加了该附件时，才使用附件 ID 删除附件。在 Outlook Web App 和 适用于设备的 OWA 中，附件 ID 仅在同一会话中有效。用户关闭外接程序时，或用户在内嵌窗体中开始撰写，然后弹出内嵌窗体以在单独的窗口中继续时，会话结束。


## 附加文件

您可以使用  **addFileAttachmentAsync** 方法在撰写窗体中将文件附加到邮件或约会，并指定文件 URI。如果文件受保护，您可以包括相应的标识或身份验证令牌作为 URI 查询字符串参数。Exchange 将向 URI 发出调用以获取附件，保护文件的 Web 服务将需要使用令牌作为进行身份验证的一种方式。

下面的 JavaScript 示例是从 Web 服务器将文件、picture.png 附加到正在撰写的邮件或约会的撰写加载项。回调方法将  **asyncResult** 作为参数，检查附加状态，并在附加成功的情况下获取附件 ID。




```js
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 附加 Outlook 项目

您可以通过指定项目的 Exchange Web Services (EWS) ID 并使用  **addItemAttachmentAsync** 方法，将 Outlook 项目（例如，电子邮件、日历或联系人项目）附加到撰写窗体中的邮件或约会。您可以通过使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法并访问 EWS 操作 [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)，来获取用户邮箱中电子邮件、日历、联系人或任务项目的 EWS ID。 [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md) 属性还提供阅读窗体中某个现有项目的 EWS ID。

以下 JavaScript 函数  `addItemAttachment` 扩展了以上第一个示例，并将项目作为附件添加到正在撰写的电子邮件或约会。此函数将要附加的项目的 EWS ID 作为实参。如果附加成功，则会获取附件 ID 以执行进一步的处理，包括在同一会话中删除该附件。




```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


 >**注释**  您可以使用撰写加载项在 Outlook Web App 或 适用于设备的 OWA 中附加定期约会的实例。但是，在 Outlook 富客户端中，尝试附加实例将导致附加定期系列（主约会）。


## 删除附件


您可以指定相应的附件 ID，并使用 [removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法在撰写窗体中从邮件或约会项目删除文件或项目附件。您应仅删除同一外接程序在同一会话中添加的附件。应确保附件 ID 与有效附件对应，否则此方法将返回错误。类似于 **addFileAttachmentAsync** 和 **addItemAttachmentAsync** 方法， **removeAttachmentAsync** 是一个异步方法。您应使用 **AsyncResult** 输出参数对象提供一个回调方法以检查状态和任何错误。还可以使用可选的 **asyncContext** 参数将任何其他参数传递给回调方法，此参数是键值对的 JSON 对象。

以下 JavaScript 函数  `removeAttachment` 将继续扩展以上示例，并从正在撰写的电子邮件或约会删除指定附件。此函数将要删除的附件的 ID 作为实参。您可以在 **addFileAttachmentAsync** 或 **addItemAttachmentAsync** 方法调用成功后获取附件 ID，并进行存储以供以后的 **removeAttachmentAsync** 方法调用使用。




```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## 添加和删除附件的提示


如果您的撰写加载项将添加和删除附件，可构造代码，以便将有效附件 ID 传递给删除附件调用，并在  **AsyncResult.error** 返回 **InvalidAttachmentId** 后处理这种情况。根据附件的位置和大小，附加文件或项目将需要一段时间来完成。以下代码包含对 **addFileAttachmentAsync**、 `write` 和 **removeAttachmentAsync** 的调用。您可能认为此调用将按顺序依次执行。


```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

尽管  **addFileAttachmentAsync** 在 **removeAttachmentAsync** 之前启动，但由于 **addFileAttachmentAsync** 为异步调用，因此 `write` 和 **removeAttachmentAsync** 调用可在 **addFileAttachmentAsync** 完成之前开始。发生这种情况时， `attachmentID` 将保持 **undefined**，您将收到一个  **removeAttachmentAsync** 调用的错误，如以下输出中所示：




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

避免此情况发生的方法是检查  `attachmentID` 是否在调用 **removeAttachmentAsync** 之前进行了定义。另一种方法是从 **addFileAttachmentAsync** 的回调方法内启动 **removeAttachmentAsync** 调用，如以下示例中所示：




```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

下面是输出的一个示例：




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

请注意， **removeAttachmentAsync** 的回调嵌套在 **addFileAttachmentAsync** 的回调中。由于 **addFileAttachmentAsync** 和 **removeAttachmentAsync** 为异步调用，因此可在 **removeAttachmentAsync** 的回调完成之前执行 **addFileAttachmentAsync** 的回调中的最后一行。


## 其他资源



- [创建适用于撰写窗体的 Outlook 外接程序](../outlook/compose-scenario.md)
    
- [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


