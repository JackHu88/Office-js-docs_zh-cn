﻿

# <a name="item"></a>item

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). item

`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|

### <a name="example"></a>示例

以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>成员

#### <a name="attachments-arrayattachmentdetailssimple-typesmdattachmentdetails"></a>attachments :Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

获取项目的附件数组。仅限阅读模式。

##### <a name="type"></a>类型:

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsrecipientsmd"></a>bcc :[Recipients](Recipients.md)

获取或设置邮件“密件抄送”行上的收件人。仅限撰写模式。

##### <a name="type"></a>类型:

*   [收件人](Recipients.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodybodymd"></a>body :[Body](Body.md)

获取一个提供用于处理项目正文的方法的对象。

##### <a name="type"></a>类型：

*   [Body](Body.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[收件人](Recipients.md)

获取或设置邮件的抄送 (Cc) 收件人。

##### <a name="read-mode"></a>阅读模式

`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。

##### <a name="compose-mode"></a>撰写模式

`cc` 属性返回 `Recipients` 对象，该对象提供用于处理邮件的**抄送**行上收件人的方法。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

获取包含特定消息的电子邮件会话的标识符。

如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。

对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

获取项目创建的日期和时间。仅限阅读模式。

##### <a name="type"></a>类型：

*   日期

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

获取项目最近一次修改的日期和时间。仅限阅读模式。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此成员。

##### <a name="type"></a>类型:

*   日期

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimetimemd"></a>end :Date|[Time](Time.md)

获取或设置约会结束的日期和时间。

将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。

##### <a name="read-mode"></a>阅读模式

`end` 属性返回 `Date` 对象。

##### <a name="compose-mode"></a>撰写模式

`end` 属性返回 `Time` 对象。

使用 [`Time.setAsync`](Time.md#setasync) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。

##### <a name="type"></a>类型：

*   Date | [Time](Time.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

以下示例通过使用 `Time` 对象的 [`setAsync`](Time.md#setasync) 方法，设置撰写模式下约会的结束时间。

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailssimple-typesmdemailaddressdetails"></a>from :[EmailAddressDetails](simple-types.md#emailaddressdetails)

获取邮件发件人的电子邮件地址。仅限阅读模式。

`from` 和 [`sender`](Office.context.mailbox.item.md#sender) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。

##### <a name="type"></a>类型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="internetmessageid-string"></a>internetMessageId :String

获取电子邮件的 Internet 消息标识符。仅限阅读模式。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。

`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。

| 类型 | 说明 | 项目类 |
| --- | --- | --- |
| 约会项目 | 这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。 | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| 邮件项目 | 这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。 | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。

`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。`itemId` 属性不等同于 Outlook 条目 ID。

`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](Office.context.mailbox.item.md#saveAsync) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](simple-types.md#asyncresult) 参数中返回项目标识符。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypeofficemailboxenumsmditemtype-string"></a>itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

获取实例表示的项的类型。

`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。

##### <a name="type"></a>类型：

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationlocationmd"></a>location :String|[Location](Location.md)

获取或设置约会的位置。

##### <a name="read-mode"></a>阅读模式

`location` 属性返回一个包含约会位置的字符串。

##### <a name="compose-mode"></a>撰写模式

`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。

##### <a name="type"></a>类型：

*   String | [Location](Location.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。

normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](Office.context.mailbox.item.md#subject) 属性。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailssimple-typesmdemailaddressdetailsrecipientsrecipientsmd"></a>optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

获取或设置可选与会者的电子邮件地址列表。

##### <a name="read-mode"></a>阅读模式

`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。

##### <a name="compose-mode"></a>撰写模式

`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailssimple-typesmdemailaddressdetails"></a>organizer :[EmailAddressDetails](simple-types.md#emailaddressdetails)

获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。

##### <a name="type"></a>类型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailssimple-typesmdemailaddressdetailsrecipientsrecipientsmd"></a>requiredAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

获取或设置必需与会者的电子邮件地址的列表。

##### <a name="read-mode"></a>阅读模式

`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。

##### <a name="compose-mode"></a>撰写模式

`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置必需与会者的方法。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="resources-emailaddressdetailssimple-typesmdemailaddressdetails"></a>resources :[EmailAddressDetails](simple-types.md#emailaddressdetails)

获取约会所需的资源。仅限阅读模式。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此成员。

##### <a name="type"></a>类型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="sender-emailaddressdetailssimple-typesmdemailaddressdetails"></a>sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

获取电子邮件发件人的电子邮件地址。仅限阅读模式。

[`from`](Office.context.mailbox.item.md#from) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。

##### <a name="type"></a>类型:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="example"></a>示例

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimetimemd"></a>start :Date|[Time](Time.md)

获取或设置约会开始的日期和时间。

将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。

##### <a name="read-mode"></a>阅读模式

`start` 属性返回 `Date` 对象。

##### <a name="compose-mode"></a>撰写模式

`start` 属性返回 `Time` 对象。

使用 [`Time.setAsync`](Time.md#setasync) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。

##### <a name="type"></a>类型：

*   Date | [Time](Time.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

以下示例通过使用 `Time` 对象的 [`setAsync`](Time.md#setasync) 方法，设置撰写模式下约会的开始时间。

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectsubjectmd"></a>subject :String|[Subject](Subject.md)

获取或设置显示在项目的主题字段中的说明。

`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。

##### <a name="read-mode"></a>阅读模式

`subject` 属性返回一个字符串。使用 [`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>撰写模式

`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>类型：

*   String | [Subject](Subject.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  to :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[收件人](Recipients.md)

获取或设置电子邮件的收件人。

##### <a name="read-mode"></a>阅读模式

`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。

##### <a name="compose-mode"></a>撰写模式

`to` 属性返回 `Recipients` 对象，该对象提供用于处理邮件的**收件人**行上收件人的方法。

##### <a name="type"></a>类型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>方法

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

将文件作为附件添加到邮件或约会。

`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。

你随后可以将该标识符与 [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`uri`| 字符串||提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。|
|`attachmentName`| 字符串||在附件上载过程中显示的附件名称。最大长度为 255 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果成功，附件标识符将在 `asyncResult.value` 属性中提供。<br/>如果上载附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>附件大小超过了允许的大小。</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>该附件的扩展名不是允许的扩展名。</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>邮件或约会具有的附件过多。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

将 Exchange 项目（如邮件）作为附件添加到邮件或约会。

`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。

你随后可以将该标识符与 [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。

如果 Office 外接程序在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`itemId`| 字符串||要附加的项目的 Exchange 标识符。最大长度为 100 个字符。|
|`attachmentName`| 字符串||要附加的项目的主题。最大长度为 255 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果成功，附件标识符将在 `asyncResult.value` 属性中提供。<br/>如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>邮件或约会具有的附件过多。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。

如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。

> **注意：**要求集 1.1 不支持 `displayReplyAllForm` 在调用中包括附件的功能。附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`formData`| 字符串 &#124; 对象|一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。<br/>**OR**<br/>包含正文或附件数据和回调函数的对象。对象定义如下：<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>字符串</td><td>&lt;可选&gt;</td><td>一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</td></tr><tr><td><code>callback</code></td><td>函数</td><td>&lt;可选&gt;</td><td>方法完成后，使用单个参数 <code>asyncResult</code>（一个 <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a> 对象）调用在 <code>callback</code> 参数中传递的函数。有关详细信息，请参阅<a href="tutorial-asynchronous.html">使用异步方法</a>。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="examples"></a>示例

以下代码将一个字符串传递到 `displayReplyAllForm` 函数。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

使用空白正文答复。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

仅使用正文答复。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

使用正文和回调答复。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。

如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。

> **注意：**要求集 1.1 不支持 `displayReplyForm` 在调用中包括附件的功能。附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`formData`| 字符串 &#124; 对象|一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。<br/>**OR**<br/>包含正文或附件数据和回调函数的对象。对象定义如下：<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>字符串</td><td>&lt;可选&gt;</td><td>一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</td></tr><tr><td><code>callback</code></td><td>函数</td><td>&lt;可选&gt;</td><td>方法完成后，使用单个参数 <code>asyncResult</code>（一个 <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a> 对象）调用在 <code>callback</code> 参数中传递的函数。有关详细信息，请参阅<a href="tutorial-asynchronous.html">使用异步方法</a>。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="examples"></a>示例

以下代码将一个字符串传递到 `displayReplyForm` 函数。

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

使用空白正文答复。

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

仅使用正文答复。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

使用正文和回调答复。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiessimple-typesmdentities"></a>getEntities() → {[Entities](simple-types.md#entities)}

获取在所选项目中找到的实体。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="returns"></a>返回：

类型：[Entities](simple-types.md#entities)

##### <a name="example"></a>示例

以下示例访问当前项目上的联系人实体。

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactsimple-typesmdcontactmeetingsuggestionsimple-typesmdmeetingsuggestionphonenumbersimple-typesmdphonenumbertasksuggestionsimple-typesmdtasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

获取所选项目中找到的指定实体类型的所有实体的数组。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.entitytype-string)|EntityType 枚举值之一。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 阅读|

##### <a name="returns"></a>返回：

如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。如果指定类型的任何实体都不存在于该项目上，该方法将返回空数组。否则，返回的数组中对象的类型取决于 `entityType` 参数中请求的实体类型。

当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。

| `entityType` 的值 | 返回的数组中对象的类型 | 所需权限级别 |
| --- | --- | --- |
| `Address` | String | **受限** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **受限** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **受限** |

类型：Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))></dd>


##### <a name="example"></a>示例

以下示例显示了如何访问表示当前项目的主题或正文中的邮政地址的字符串数组。

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactsimple-typesmdcontactmeetingsuggestionsimple-typesmdmeetingsuggestionphonenumbersimple-typesmdphonenumbertasksuggestionsimple-typesmdtasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。


  `getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) 规则元素中定义的正则表达式的实体。

##### <a name="parameters"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="returns"></a>返回：

如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。


类型：Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>


#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。

例如，考虑一个外接程序清单具有以下 `Rule` 元素：

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](Body.md#getAsync) 方法检索整个正文。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="returns"></a>返回：

一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。

<dl class="param-type">

<dt>
类型</dt>


<dd>对象</dd>

</dl>

##### <a name="example"></a>示例

以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-arraystring"></a>getRegExMatchesByName(name) → (nullable) {Array.<String>}

返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。

> **注意：**在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。

如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。

##### <a name="parameters"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|

##### <a name="returns"></a>返回：

一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。

<dl class="param-type">

<dt>类型</dt>

<dd>数组。<String></dd>

</dl>

##### <a name="example"></a>示例

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

以异步方式返回邮件的主题或正文中选定的数据。

如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。|
|`options`| 对象| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="returns"></a>返回：

作为字符串的所选数据的格式由 `coercionType` 确定。

<dl class="param-type">

<dt>
类型</dt>


<dd>字符串</dd>

</dl>

##### <a name="example"></a>示例

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

异步加载所选项目上此外接程序的自定义属性。

自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| 函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](CustomProperties.md) 对象提供。此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。| |`userContext`| 对象| &lt;可选&gt;|开发人员可以提供他们想要在回调函数中访问的任何对象。此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

将附件从邮件或约会中删除。

`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`attachmentId`| 字符串||要删除的附件的标识符。字符串的最大长度为 100 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>附件标识符不存在。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

以下代码删除包含标识符 '0' 的附件。

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
