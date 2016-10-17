
# <a name="get,-set,-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>在 Outlook 中撰写约会或邮件时获取、设置或添加收件人


适用于 Office 的 JavaScript API 提供在约会或邮件的撰写窗体中分别获取、设置或添加收件人的异步方法（[Recipients.getAsync](../../reference/outlook/Recipients.md)、[Recipients.setAsync](../../reference/outlook/Recipients.md) 或 [Recipients.addAysnc](../../reference/outlook/Recipients.md)）。这些异步方法仅对撰写外接程序可用。若要使用这些方法，请确保已为 Outlook 正确设置外接程序清单，以在撰写窗体中激活外接程序，如[创建适用于撰写窗体的 Outlook 外接程序](../outlook/compose-scenario.md)中所述。

部分表示约会或邮件中的收件人的属性在撰写窗体和阅读窗体中可以进行阅读访问。这些属性包括约会的 [optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) 和 [requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)，以及邮件的 [cc](../../reference/outlook/Office.context.mailbox.item.md) 和 [to](../../reference/outlook/Office.context.mailbox.item.md)。在阅读窗体中，您可以直接从父对象访问属性，例如：




```js
item.cc
```

但是在撰写窗体中，由于用户和外接程序可能正在同时插入或更改收件人，所以必须使用异步方法  **getAsync** 获取这些属性，如以下示例所示：




```js
item.cc.getAsync
```

这些属性只在撰写窗体（而非阅读窗体）中可进行写入访问。

与适用于 Office 的 JavaScript API 中的大多数异步方法一样，**getAsync**、**setAsync** 和 **addAsync** 采用可选输入参数。有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)中的[向异步方法传递可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md)。


## <a name="to-get-recipients"></a>获取收件人


此部分显示的代码示例用于获取正在撰写的约会或邮件的收件人，并显示收件人的电子邮件地址。代码示例假设外接程序清单中有在撰写窗体中为约会或邮件激活外接程序的规则，如下所示。 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

在适用于 Office 的 JavaScript API 中，由于代表约会收件人的属性（ **optionalAttendees** 和 **requiredAttendees**）与代表邮件收件人的属性（ [bcc](../../reference/outlook/Office.context.mailbox.item.md)、 **cc** 和 **to**）不同，所以应首先使用 [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) 属性确定正在撰写的项目是约会还是邮件。在撰写模式中，所有这些约会和邮件属性都是 [Recipients](../../reference/outlook/Recipients.md) 对象，所以您可以应用异步方法 **Recipients.getAsync** 获取相应的收件人。 

若要使用  **getAsync**，请提供回调方法检查异步  **getAsync** 调用返回的状态、结果和任何错误。您可以使用可选 _asyncContext_ 形参为回调方法提供任意实参。回调方法会返回 _asyncResult_ 输出形参。您可以使用 **status** 形成对象的 **error** 和 [error](../../reference/outlook/simple-types.md) 属性检查异步调用的状态和任何错误消息，以及使用 **value** 属性获取实际收件人。收件人以 [EmailAddressDetails](../../reference/outlook/simple-types.md) 对象数组的形式表示。

请注意，由于  **getAsync** 方法是异步方法，如果根据成功获取收件人后还有后续操作，则异步调用成功完成时，应只在相应回调方法中组织代码启动这类操作。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-set-recipients"></a>设置收件人


此部分显示的代码示例会设置用户正在撰写的约会或邮件的收件人。设置收件人将覆盖现有的全部收件人。与之前获取撰写窗体中收件人的示例相似，此示例假设已在撰写窗体中为约会和邮件激活外接程序。此示例首先验证撰写的项目是约会还是邮件，以便对代表约会或邮件收件人的合适属性应用异步方法 **Recipients.setAsync**。

调用  **setAsync** 时，请提供一个数组作为 _recipients_ 形参的输入实参，采用以下任一格式提供：


- 为 SMTP 地址的字符串数组。
    
- 字典的数组，每个字典都包含显示名称和电子邮件地址，如下面的代码示例中所示。
    
- **EmailAddressDetails** 对象的数组，与 **getAsync** 方法返回的数组相似。
    
您还可以选择提供一个回调方法作为  **setAsync** 方法的输入实参，以确保基于成功设置收件人的任何代码只在成功时才会执行。还可以为使用可选 _asyncContext_ 形参的回调方法提供任意实参。如果使用回调方法，可以访问 _asyncResult_ 输出形参，并使用 **AsyncResult** 形参对象的 **status** 和 **error** 属性检查异步调用的状态和所有错误消息。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="to-add-recipients"></a>添加收件人


如果不想覆盖约会或邮件中的任何现有收件人，可以使用  **Recipients.addAsync** 异步方法追加收件人，而不是使用 **Recipients.setAsync**。 **addAsync** 工作原理与 **setAsync** 相似，因为也需要 _recipients_ 输入实参。还可以选择使用 asyncContext 形参为回调提供回调方法和任意实参。然后，可以使用回调方法的 **asyncResult** 输出形参检查异步 _addAsync_ 调用的状态、结果和任何错误。以下示例检查正在撰写的项目是否是约会，并为该约会追加两个必需参与者。


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="additional-resources"></a>其他资源



- [在 Outlook 的撰写窗体中获取和设置项目数据](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](../outlook/item-data.md)
    
- [创建适用于撰写窗体的 Outlook 外接程序](../outlook/compose-scenario.md)
    
- [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [在 Outlook 中撰写约会或邮件时获取或设置主题](../outlook/get-or-set-the-subject.md)
    
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](../outlook/insert-data-in-the-body.md)
    
- [在 Outlook 中撰写约会时获取或设置位置](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [在 Outlook 中撰写约会时获取或设置时间](../outlook/get-or-set-the-time-of-an-appointment.md)
    
