

# NotificationMessages

## NotificationMessages

`NotificationMessages` 对象作为项目的 [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) 属性返回。

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

### 方法

####  addAsync(key, JSONmessage, [options], [callback])

向项目添加通知。

每封邮件中最多有 5 个通知。设置过多的通知将返回 `NumberOfNotificationMessagesExceeded` 错误。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`key`| String||用于引用此通知邮件的开发人员指定的项。开发人员可以在以后用它来修改此邮件。其长度不能超过 32 个字符。|
|`JSONmessage`| Object||一个包含要添加到项目的通知邮件的 JSON 对象。它包含下列属性。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>说明</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>指定邮件的类型。如果类型是 <code>ProgressIndicator</code> 或 <code>ErrorMessage</code>，则自动提供一个图标，而且邮件不是持久性的。因此，图标和持久性的属性对于这些类型的邮件都是无效的。包括它们将导致 <code>ArgumentException</code>。如果类型是 <code>ProgressIndicator</code>，则在操作完成时开发人员应删除或替换进度指示器。</td></tr><tr><td><code>icon</code></td><td>字符串</td><td>对在清单的 <code>Resource</code> 部分中定义的图标的引用。它将显示在信息栏区域。仅当类型是 <code>InformationalMessage</code> 时才适用。为不受支持的类型指定此参数将导致异常。</td></tr><tr><td><code>message</code></td><td>String</td><td>通知邮件的文本。最大长度为 150 个字符。如果开发人员传入更长的字符串，则会引发 <code>ArgumentOutOfRange</code> 异常。</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>仅当类型是 <code>InformationalMessage</code> 时才适用。如果为 <code>true</code>，则保留邮件，直到此外接程序或用户删除该邮件。如果为 <code>false</code>，则在用户导航到其他项目时删除该邮件。对于错误通知，邮件将一直保留，直到用户看过一次。为不受支持的类型指定此参数将引发异常。</td></tr></tbody></table>|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 |

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### 示例

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", {
  type: "progressIndicator",
  message : "An add-in is processing this message."
});
Office.context.mailbox.item.notificationMessages.addAsync("information", {
  type: "informationalMessage",
  message : "The add-in processed this message.",
  icon : "iconid",
  persistent: false
});
Office.context.mailbox.item.notificationMessages.addAsync("error", {
  type: "errorMessage",
  message : "The add-in failed to process this message."
});
```

####  getAllAsync([options], [callback])

返回某个项目的所有项和邮件。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

在成功完成时，`asyncResult.value` 属性将包含一个组 [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails) 对象。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### 示例

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", {
      type: "informationalMessage",
      message : "Found " + asyncResult.value.length + " notifications.",
      icon : "iconid",
      persistent: false
    });
  }
});
```

####  removeAsync(key, [options], [callback])

删除某个项目的通知邮件。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`key`| String||要删除的通知邮件的项。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

如果找不到该项，则在 `KeyNotFound` 属性中返回 `asyncResult.error` 错误。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### 示例

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  replaceAsync(key, JSONmessage, [options], [callback])

将带有给定项的通知邮件替换为另一封邮件。

如果带有指定项的通知邮件不存在，`replaceAsync` 将添加通知。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`key`| String||要替换的通知邮件的项。长度不能超过 32 个字符。|
|`JSONmessage`| Object||一个包含要替换现有邮件的新通知邮件的 JSON 对象。它包含下列属性。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>说明</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>指定邮件的类型。如果类型是 <code>ProgressIndicator</code> 或 <code>ErrorMessage</code>，则自动提供一个图标，而且邮件不是持久性的。因此，图标和持久性的属性对于这些类型的邮件都是无效的。包括它们将导致 <code>ArgumentException</code>。如果类型是 <code>ProgressIndicator</code>，则在操作完成时开发人员应删除或替换进度指示器。</td></tr><tr><td><code>icon</code></td><td>字符串</td><td>对在清单的 <code>Resource</code> 部分中定义的图标的引用。它将显示在信息栏区域。仅当类型是 <code>InformationalMessage</code> 时才适用。为不受支持的类型指定此参数将导致异常。</td></tr><tr><td><code>message</code></td><td>String</td><td>通知邮件的文本。最大长度为 150 个字符。如果开发人员传入更长的字符串，则会引发 <code>ArgumentOutOfRange</code> 异常。</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>仅当类型是 <code>InformationalMessage</code> 时才适用。如果为 <code>true</code>，则保留邮件，直到此外接程序或用户删除该邮件。如果为 <code>false</code>，则在用户导航到其他项目时删除该邮件。对于错误通知，邮件将一直保留，直到用户看过一次。为不受支持的类型指定此参数将引发异常。</td></tr></tbody></table>|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 |

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### 示例

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
  type: "informationalMessage",
  message : "The message was processed successfully.",
  icon : "iconid",
  persistent: false
});
```
