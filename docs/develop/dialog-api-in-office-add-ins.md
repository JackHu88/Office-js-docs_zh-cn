# <a name="use-the-dialog-api-in-your-office-add-ins"></a>在 Office 外接程序中使用对话框 API 

可以在 Office 外接程序中使用[对话框 API](../../reference/shared/officeui.md) 打开对话框。本文提供了有关如何在 Office 外接程序中使用对话框 API 的指南。

> **注意：**若要了解对话框 API 目前的受支持情况，请参阅[对话框 API 要求集](../../reference/requirement-sets/dialog-api-requirement-sets.md)。目前，Word、Excel、PowerPoint 和 Outlook 支持对话框 API。

不妨从任务窗格或内容外接程序/[外接程序命令](https://dev.office.com/docs/add-ins/design/add-in-commands)打开对话框，从而： 

- 显示无法直接在任务窗格中打开的登录页。
- 为外接程序中的某些任务提供更多屏幕空间，或甚至整个屏幕。
- 托管在任务窗格中显得太小的视频。

>**注意：**由于重叠 UI 可能会令用户生厌，因此除非应用场景需要，否则不要从任务窗格打开对话框。在考虑如何使用任务窗格区域时，请注意任务窗格可以带有选项卡。有关示例，请参阅 [Excel 外接程序 JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。

下图为一个对话框示例。 

![外接程序命令](../../images/Auth0DialogOpen.PNG)

请注意，对话框总是在屏幕的中心打开。用户可以移动并调整其大小。对话框是*非模式*窗口。也就是说，用户可以继续同时与主机 Office 应用程序中的文档以及任务窗格中的主机页（若有）进行交互。

## <a name="dialog-api-scenarios"></a>对话框 API 应用场景

Office JavaScript API 支持以下应用场景，其在 [Office.context.ui 命名空间](../../reference/shared/officeui.md)中使用 [Dialog](../../reference/shared/officeui.dialog.md) 对象和两个函数。 

### <a name="opening-a-dialog-box"></a>打开对话框

为了打开对话框，任务窗格中的代码会调用 [displayDialogAsync](../../reference/shared/officeui.displaydialogasync.md) 方法，然后将应打开的资源 URL 传递给该方法。这通常是一个页面，但它可能是 MVC 应用程序中的控制器方法、路由、Web 服务方法或任何其他资源。在本文中，“页面”或“网页”指对话框中的资源。下面展示了一个非常简单的示例。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html'); 
```

> **注意：**

> - URL 使用 HTTP**S** 协议。对话框中加载的所有页面都必须要遵循此要求，而不仅仅是加载的第一个页面。
> - 域与主机页的域相同，主机页可以是任务窗格中的页面，也可以是外接程序命令的[函数文件](https://dev.office.com/reference/add-ins/manifest/functionfile)。这要求：传递到 `displayDialogAsync` 方法的页面、控制器方法或其他资源必须与主机页位于相同的域。 

在第一个页面（或其他资源）加载后，用户可以转到使用 HTTPS 的任意网站（或其他资源）。还可以将第一个页面设计为直接重定向到另一个站点。 

默认情况下，对话框的高度和宽度占设备屏幕的 80%。不过，你也可以设置不同的百分比，只需将配置对象传递给方法即可，如下面的示例所示。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20}); 
```

有关实现这一点的示例外接程序，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

将两个值均设置为 100% 可有效提供全屏体验。（有效最大值为 99.5%，窗口仍可移动和调整大小。）

>**注意：**只能从主机窗口打开一个对话框。尝试打开另一个对话框会生成错误。（有关详细信息，请参阅 [displayDialogAsync 返回的错误](#errors-from-displaydialogAsync)。）比方说，如果用户从任务窗格打开对话框，则无法从任务窗格中的其他页面打开第二个对话框。不过，如果是从[外接程序命令](https://dev.office.com/docs/add-ins/design/add-in-commands)打开对话框，那么只要选择此命令，就会打开一个新的（但不可见的）HTML 文件。这会新建一个（不可见的）主机窗口，所以每个这样的窗口都可以启动自己的对话框。 

### <a name="take-advantage-of-a-performance-option-in-office-online"></a>使用 Office Online 中的性能选项

`displayInIframe` 属性是可以传递到 `displayDialogAsync` 的配置对象中的附加属性。当将此属性设置为 `true` 且外接程序在 Office Online 打开的文档中运行时，对话框将作为浮动 iframe 而非独立窗口打开，这样可以使对话框打开速度更快。示例如下。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true}); 
```

默认值为 `false`，与完全遗漏此属性时完全相同。

如果外接程序未在 Office Online 中运行，则会忽略 `displayInIframe`，但是对其存在没有危害。

> **注意：**如果对话框将随时重定向至 iframe 无法打开的页面，则***不***应使用 `displayInIframe: true`。例如，许多热门 Web 服务的登录页（如 Google 和 Microsoft 帐户）都无法在 iframe 中打开。 

### <a name="sending-information-from-the-dialog-box-to-the-host-page"></a>将信息从对话框发送到主机页

对话框无法与任务窗格中的主机页进行通信，除非：

- 对话框中的当前页面与主机页在同一个域中。
- Office JavaScript 库已在页面中加载。（与使用 Office JavaScript 库的所有页面一样，页面脚本必须为 `Office.initialize` 属性分配方法，尽管方法可以是空的。有关详细信息，请参阅[初始化外接程序](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)。） 

对话框页中的代码使用 `messageParent` 函数向主机页发送布尔值或字符串消息。字符串可以是单词、句子、XML blob、字符串化 JSON 或其他任何能够序列化成字符串的内容。示例如下。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true); 
}
```

>**注意：** 

> - *只有*两个 Office API 可以在对话框中调用，`messageParent` 函数就是其中一个。（另一个是 `Office.context.requirements.isSetSupported`。有关详细信息，请参阅[指定 Office 主机和 API 要求](https://github.com/OfficeDev/office-js-docs/blob/master/docs/overview/specify-office-hosts-and-api-requirements.md)。）
> - `messageParent` 函数只能在与主机页位于同一域（包括协议和端口）的页面上调用。

在下一个示例中，`googleProfile` 是用户 Google 个人资料的字符串化版本。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile); 
}
```

必须将主机页配置为接收消息。为此，可以向 `displayDialogAsync` 的原始调用添加回叫参数。回叫会为 `DialogMessageReceived` 事件分配处理程序。示例如下。关于此代码，请注意以下几点：

- Office 将 [AsyncResult ](https://dev.office.com/reference/add-ins/shared/asyncresult) 对象传递给回叫。表示尝试打开对话框的结果，不表示对话框中任何事件的结果。若要详细了解此区别，请参阅[处理错误和事件](#handling-errors-and-events)部分。 
- `asyncResult` 的 `value` 属性设置为 [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) 对象，该对象位于主机页（而不是对话框的执行上下文）中。
- `processMessage` 是用于处理事件的函数。可以根据需要任意命名。 
- `dialog` 变量的声明范围比回叫更广，因为 `processMessage` 中也会引用该变量。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
); 
```

下面是一个非常简单的示例，展示了 `DialogMessageReceived` 事件的处理程序。关于此代码，请注意以下几点：

- Office 将 `arg` 对象传递给处理程序。它的 `message` 属性是对话框中的 `messageParent` 调用发送的布尔值或字符串。在此示例中，它是 Microsoft 帐户或 Google 等服务的用户配置文件的字符串化表示。因此，使用 `JSON.parse` 将其反序列化回对象。
- 未显示 `showUserName` 实现。它可能在任务窗格上显示定制的欢迎消息。

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

在用户完成与对话框的交互后，消息处理程序应关闭对话框，如以下示例所示。关于此代码，请注意以下几点：

- `dialog` 对象必须是 `displayDialogAsync` 调用返回的对象。 
- `dialog.close` 调用指示 Office 立即关闭对话框。

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

有关使用这些技术的示例外接程序，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

如果外接程序在收到消息后需要打开任务窗格的其他页面，可以使用 `window.location.replace` 方法（或 `window.location.href`）作为处理程序的最后一行。示例如下。

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

有关实现这一点的外接程序示例，请参阅[在 PowerPoint 外接程序中使用 Microsoft Graph 插入 Excel 图表](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)示例。 

#### <a name="conditional-messaging"></a>条件消息

因为你可以从对话框发送多个 `messageParent` 调用，但是在 `DialogMessageReceived` 事件的主机页中只有一个处理程序，所以处理程序不得不使用条件逻辑来区分不同的消息。例如，如果对话框提示用户登录标识提供程序（如 Microsoft 帐户或 Google），则会以消息形式发送用户配置文件。如果身份验证失败，对话框应将错误消息发送给主机页，如以下示例所示。关于此代码，请注意以下几点：

- `loginSuccess` 变量通过读取标识提供程序的 HTTP 响应进行初始化。
- 未显示 `getProfile` 和 `getError` 函数的实现。这两个函数均从查询参数或 HTTP 响应的正文获取数据。
- 根据登录是否成功，发送不同类型的匿名对象。两者都有 `messageType` 属性。不同之处在于，一个有 `profile` 属性，另一个有 `error` 属性。

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage); 
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage); 
}
```

有关使用条件消息的示例，请参阅 

- [使用 Auth0 服务简化社交登录的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [使用 OAuth.io 服务简化热门联机服务访问的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

主机页中的处理程序代码使用 `messageType` 属性的值设置分支，如以下示例所示。请注意，`showUserName` 函数的用法与上面的示例相同，`showNotification` 函数在主机页的 UI 中显示错误。 

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

### <a name="closing-the-dialog-box"></a>关闭对话框

可以在对话框中实现一个用于关闭对话框的按钮。为此，该按钮的单击事件处理程序应使用 `messageParent` 通知主机页该按钮已获得单击。示例如下。

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage); 
}
``` 

`DialogMessageReceived` 的主机页处理程序将调用 `dialog.close`，如以下示例所示。（请参阅本文前面的示例，其中展示了对话框对象的初始化方式。）


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

有关使用此技术的示例，请参阅 [Office 外接程序的用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)存储库中的[对话框导航设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)。

即使你没有自己的关闭对话框 UI，最终用户也可以通过选择右上角的 **X** 关闭对话框。此操作将触发 `DialogEventReceived` 事件。如果主机窗格需要知道此事件何时发生，应为此事件声明一个处理程序。有关详细信息，请参阅[对话框窗口中的错误和事件](#errors-and-events-in-the-dialog-window)部分。

## <a name="handling-errors-and-events"></a>处理错误和事件 

代码应处理两类事件：

- `displayDialogAsync` 调用返回的错误，因为无法创建对话框。 
- 对话框窗口中的错误和其他事件。

### <a name="errors-from-displaydialogasync"></a>DisplayDialogAsync 返回的错误

除常规的平台和系统错误外，调用 `displayDialogAsync` 会返回以下三个特定错误。

|代码编号|含义|
|:-----|:-----|
|12004|传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。|
|12005|传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。需要使用 HTTPS。（在 Office 的某些版本中，返回 12005 的错误消息与返回 12004 错误消息是相同的。）|
|12007|已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。|

调用 `displayDialogAsync` 时，总是将 [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) 对象传递给它的回叫函数。如果调用成功（即对话框窗口已打开），`AsyncResult` 对象的 `value` 属性是 [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) 对象。有关示例，请参阅[将信息从对话框发送到主机页](#sending-information-from-the-dialog-to-the-host-page)部分。如果调用 `displayDialogAsync` 失败，不会创建窗口，`AsyncResult` 对象的 `status` 属性设置为“failed”，并且会填充对象的 `error` 属性。应始终有用于测试 `status` 并在出错时进行响应的回叫。无论代码编号是什么，下面的示例仅报告错误消息。 

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', 
function (asyncResult) {
    if (asyncResult.status === "failed") { 
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message); 
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
}); 
```

### <a name="errors-and-events-in-the-dialog-window"></a>对话框窗口中的错误和事件

对话框中的三个错误和事件（具有代码编码）会在主机页中触发 `DialogEventReceived` 事件。 

|代码编号|含义|
|:-----|:-----|
|12002|下列一种含义：<br> - 传递给 `displayDialogAsync` 的 URL 没有对应的页面。<br> - 传递给 `displayDialogAsync` 的页面已加载，但对话框定向到找不到或无法加载的页面，或者已定向到使用无效语法的 URL。|
|12003|对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。|
|12006|对话框已关闭，通常是因为用户选择了 **X** 按钮。|

代码可以在调用 `displayDialogAsync` 时为 `DialogEventReceived` 事件分配处理程序。下面展示了一个非常简单的示例。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', 
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
); 
```

下面的示例展示了 `DialogEventReceived` 事件的处理程序，其为每个错误代码创建自定义错误消息。 

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

有关这样处理错误的示例外接程序，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

 
## <a name="passing-information-to-the-dialog-box"></a>向对话框传递信息

有时，主机页需要向对话框传递信息。完成此操作的方式主要分为两种：

- 向传递给 `displayDialogAsync` 的 URL 添加查询参数。 
- 将信息存储在主机窗口和对话框都可访问的位置。这两个窗口不共享通用会话存储，但*如果它们具有相同的域*（包括端口号，若有），则共享通用[本地存储](http://www.w3schools.com/html/html5_webstorage.asp)。

### <a name="using-local-storage"></a>使用本地存储

为了使用本地存储，代码在调用 `displayDialogAsync` 之前在主机页中调用 `window.localStorage` 对象的 `setItem` 方法，如以下示例所示。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

对话框窗口中的代码在需要时读取项目，如以下示例所示。

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

有关以此种方式使用本地存储的外接程序的示例，请参阅 

- [使用 Auth0 服务简化社交登录的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [使用 OAuth.io 服务简化热门联机服务访问的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="using-query-parameters"></a>使用查询参数

下面的示例展示了如何使用查询参数传递数据。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248'); 
```

有关使用此技术的示例，请参阅[在 PowerPoint 外接程序中使用 Microsoft Graph 插入 Excel 图表](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)。

对话框窗口中的代码可以解析 URL 并读取参数值。

> **注意**：Office 会自动向传递给 `displayDialogAsync` 的 URL 添加查询参数 `_host_info`。（附加在自定义查询参数（若有）之后，不会附加到对话框导航到的任何后续 URL。）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。将相同的值添加到对话框的会话存储中。同样，*代码不得读取此值，也不得写入此值*。

## <a name="using-the-dialog-apis-to-show-a-video"></a>使用对话框 API 显示视频

若要在对话框中显示视频，请执行以下操作：

1.  创建内容仅有 iframe 的页面。iframe 的 `src` 属性指向联机视频。视频 URL 必须使用 HTTP**S** 协议。在本文中，我们将此页面称为“video.dialogbox.html”。下面是一个标注示例。

        <iframe class="ms-firstrun-video__player"  width="640" height="360" 
            src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1" 
            frameborder="0" allowfullscreen>
        </iframe>

2.  video.dialogbox.html 页面必须与主机页位于相同的域。
3.  在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。
4.  如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。有关详细信息，请参阅[对话框窗口中的错误和事件](#errors-and-events-in-the-dialog-window)部分。

有关在对话框中显示视频的示例，请参阅 [Office 外接程序的用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)存储库中的[视频展示位置设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)。

![在外接程序对话框中显示的视频的屏幕截图。](../../images/VideoPlacematDialogOpen.PNG)

## <a name="using-the-dialog-apis-in-an-authentication-flow"></a>在身份验证流中使用对话框 API

对话框 API 的主要应用场景是为不允许在 Iframe 中打开登录页的资源或标识提供程序（如 Microsoft 帐户、Office 365、Google 和 Facebook）启用身份验证。 

>**注意：**将对话框 API 用于此方案时，请*不*要在 `displayDialogAsync` 的调用中使用 `displayInIframe: true` 选项。请参阅此文章上文部分了解有关此选项的详细信息。 

下面是一个简单的典型身份验证流。 

1. 对话框中打开的第一个页面是外接程序的域（即主机窗口的域）中托管的本地页面（或其他资源）。此页面可以显示简单的 UI，提示用户“请稍候，我们正在将你重定向到可以登录 *NAME-OF-PROVIDER* 的页面。”此页面中的代码使用传递给对话框的信息构建标识提供程序的登录页 URL，如[向对话框传递信息](#passing-information-to-the-dialog-box)中所述。 
2. 然后，对话框窗口重定向到登录页。URL 包含一个查询参数，用于提示标识提供程序在用户登录特定页面后重定向对话框窗口。在本文中，我们将此页面称为 "redirectPage.html"。（*此页面必须与主机窗口位于相同域中*，因为对话框窗口传递登录尝试结果的唯一方法就是调用 `messageParent`，而它只能在与主机窗口位于同一域的页面上调用）。 
2. 标识提供程序的服务处理来自对话框窗口的传入 GET 请求。如果用户已经登录，它会立即将窗口重定向到 redirectPage.html，并将用户数据作为查询参数添加。如果用户尚未登录，提供程序的登录页会显示在窗口中，以便用户登录。对于大多数提供程序，如果用户无法成功登录，提供程序会在对话框窗口中显示错误页面，而不会重定向到 redirectPage.html。用户必须通过选择右上角的 **X** 来关闭窗口。如果用户成功登录，则对话框窗口会重定向到 redirectPage.html，并且用户数据会作为查询参数添加。
3. 当 redirectPage.html 页面打开时，它会调用 `messageParent` 向主机页报告登录是否成功，而且还会视情况报告用户数据或错误数据。 
4. `DialogMessageReceived` 事件在主机页中触发，其处理程序关闭对话框窗口，并视情况对消息进行其他处理。 

有关使用此模式的示例外接程序，请参阅：

- [在 PowerPoint 外接程序中使用 Microsoft Graph 插入 Excel 图表](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)：对话框窗口最初打开的资源是没有自己视图的控制器方法。然后，其重定向到 Office 365 登录页。
- [Office 外接程序 Office 365 客户端 AngularJS 身份验证](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)：对话框窗口最初打开的资源是一个页面。 

#### <a name="supporting-multiple-identity-providers"></a>支持多个标识提供程序

如果外接程序允许用户选择提供程序（如 Microsoft 帐户、Google 或 Facebook），你需要使用本地第一个页面（见前一部分），为用户提供用于选择提供程序的 UI。用户的选择会触发登录 URL 的构建并重定向到该 URL。 

有关使用此模式的示例，请参阅[使用 Auth0 服务简化社交登录的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-Auth0)。

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>在外接程序中授权外部资源

在现代网络中，Web 应用程序是安全主体（就像用户一样），拥有自己的标识以及对联机资源（如 Office 365、Google Plus、Facebook 或 LinkedIn）的权限。在部署前，需要先向资源提供程序注册应用程序。注册内容包括： 

- 应用程序访问用户资源所需的权限的列表。
- 当应用程序访问服务时资源服务应返回访问令牌的 URL。  

如果用户在应用程序中调用访问资源服务中用户数据的函数，系统会提示用户登录相应的服务，然后提示用户授予应用程序访问用户资源所需的权限。然后，服务将登录窗口重定向到先前注册的 URL，并传递访问令牌。应用程序使用访问令牌访问用户资源。 

你可以使用对话框 API 来管理此过程，方法是使用与用户登录流类似的流，或使用[处理慢速网络](#addressing-a-slow-network)中介绍的其他流。唯一的区别是：

- 如果用户先前未向应用程序授予所需的权限，则登录后会在对话框中看到这样做的提示。 
- 对话框窗口使用 `messageParent` 发送字符串化访问令牌，或将访问令牌存储在主机窗口可以检索到的位置，从而将访问令牌发送给主机窗口。令牌具有时间限制，但在持续期间，主机窗口可以使用它直接访问用户资源，而无需进一步提示。

下面的示例使用对话框 API 实现此目的：

- [在 PowerPoint 外接程序中使用 Microsoft Graph 插入 Excel 图表](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - 将访问令牌存储在数据库中。
- [使用 OAuth.io 服务简化热门联机服务访问的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

#### <a name="more-information-about-authentication-and-authorization-in-add-ins"></a>有关外接程序中身份验证和授权的详细信息

- [在 Office 外接程序中授权外部服务](https://dev.office.com/docs/add-ins/develop/auth-external-add-ins)
- [Office JavaScript API 帮助程序库](https://github.com/OfficeDev/office-js-helpers) 


## <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>将 Office 对话框 API 与单页应用程序和客户端路由结合使用

如果外接程序使用客户端路由（单页应用程序通常这样做），则可以选择将路由 URL 传递给 [ displayDialogAsync ](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) 方法，而不是传递各个完整 HTML 页面的 URL。 

> **重要说明：**对话框位于新窗口中，其中包含它自己的执行上下文。如果你传递路由，则基本页及其所有初始化和引导代码会在这个新的上下文中再次运行，且所有变量都会在对话框中设置为各自的初始值。因此，此技术会在对话框窗口中启动应用程序的第二个实例。在对话框窗口中更改变量的代码不会更改相同变量的任务窗格版本。同样，对话框窗口有其自己的会话存储，任务窗格中的代码无法访问此类存储。 

