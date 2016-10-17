

# <a name="simple-types"></a>简单类型

####  <a name="asyncresult"></a>AsyncResult

用于封装异步请求的结果的对象，包括状态和错误信息（如果请求失败）。

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`asyncContext`| Object|获取与传入时状态相同的传递给调用方法的可选 `asyncContext` 参数的对象。|
|`error`| 错误|如果出现任何错误，获取提供错误描述的 Error 对象。|
|`status`| [Office.AsyncResultStatus](Office.md#.asyncresultstatus-string)|获取异步操作的状态。|
|`value`| Object|获取此异步操作的负载或内容（如有）。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
#### <a name="attachmentdetails"></a>AttachmentDetails

表示服务器中一个项目上的附件。仅限阅读模式。

`AttachmentDetail` 对象的数组作为 `attachments` 或 `Appointment` 对象的 `Message` 属性返回。

##### <a name="properties:"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`attachmentType`| [Office.MailboxEnums.AttachmentType](Office.MailboxEnums.md#attachmenttype-string)|获取一个指示附件类型的值。|
|`contentType`| 字符串|获取附件的 MIME 内容类型。|
|`id`| String|获取附件的 Exchange 附件 ID。|
|`isInline`| Boolean|获取指示是否应在项目正文中显示附件的值。|
|`name`| String|获取附件的名称。|
|`size`| 数字|获取以字节为单位的附件大小。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="contact"></a>Contact

表示存储在服务器上的联系人。仅限阅读模式。

与电子邮件或约会关联的联系人列表在由活动项的 `contacts` 或 `Entities` 方法返回的 [`getEntities`](simple-types.md#entities) 对象的 `getEntitiesByType` 属性中返回。

##### <a name="properties:"></a>属性：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;可为 null&gt;|包含与联系人关联的邮件和街道地址的字符串数组。|
|`businessName`| 字符串| &lt;可为 null&gt;|包含与联系人关联的企业名称的字符串。|
|`emailAddresses`| Array.&lt;String&gt;| &lt;可为 null&gt;|包含与联系人关联的 SMTP 电子邮件地址的字符串数组。|
|`personName`| String| &lt;可为 null&gt;|包含与联系人关联的人员姓名的字符串。|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;可为 null&gt;|包含与联系人关联的每个电话号码的 `PhoneNumber` 对象的数组。|
|`urls`| Array.&lt;String&gt;| &lt;可为 null&gt;|包含与联系人关联的 Internet URL 的字符串数组。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 阅读|
####  <a name="emailaddressdetails"></a>EmailAddressDetails

提供电子邮件或约会的发件人或指定收件人的电子邮件属性。

##### <a name="type:"></a>类型：

*   对象

##### <a name="properties:"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`appointmentResponse`| [Office.MailboxEnums.ResponseType](Office.MailboxEnums.md#responsetype-string)|获取参与者返回的约会响应。此属性仅适用于约会的参与者，由 [`optionalAttendees`](Office.context.mailbox.item.md#optionalattendees-arrayemailaddressdetailsrecipients) 或 [`requiredAttendees`](Office.context.mailbox.item.md#requiredattendees-arrayemailaddressdetailsrecipients) 属性表示。在其他方案中，此属性将返回 `undefined`。|
|`displayName`| String|获取与电子邮件地址关联的显示名称。|
|`emailAddress`| String|获取 SMTP 电子邮件地址。|
|`recipientType`| [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#recipienttype-string)|获取收件人的电子邮件地址类型。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
#### <a name="emailuser"></a>EmailUser

表示 Exchange Server 上的电子邮件帐户。

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`displayName`| 字符串|获取与电子邮件地址关联的显示名称。|
|`emailAddress`| String|获取 SMTP 电子邮件地址。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="entities"></a>Entities

表示电子邮件或约会中找到的实体集合。仅限阅读模式。

`Entities` 对象是项目（电子邮件或约会）包含一个或多个服务器找到的实体时由 `getEntities` 和 `getEntitiesByType` 方法返回的实体数组的容器。可以使用代码中的这些实体为查看器提供附加上下文信息，如对项目中找到的地址的映射或打开项目中找到的电话号码的拨号程序。

如果项目中不存在属性中指定类型的实体，则与该实体关联的属性为 `null`。例如，如果消息包含街道地址和电话号码，`addresses` 属性和 `phoneNumbers` 属性将包含信息，其他属性将为 `null`。

若要被识别为地址，字符串必须包含至少具有街道编号、街道名称、城市、州和邮政编码等元素的子集的美国通讯地址。

若要被识别为电话号码，字符串必须包含北美电话号码格式。

实体识别有赖于基于计算机了解大量数据的自然语言识别。实体的识别是不确定的，其成功有时取决于项中的特定上下文。

属性数组由 `getEntitiesByType` 方法返回时，仅指定实体的属性包含数据；其他所有属性均为 `null`。

##### <a name="properties:"></a>属性：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;可为 null&gt;|获取在电子邮件或约会中找到的物理地址（街道或邮寄地址）。|
|`contacts`| Array.&lt;[Contact](simple-types.md#contact)&gt;| &lt;可为 null&gt;|获取电子邮件地址或约会中找到的联系人。|
|`emailAddresses`| Array.&lt;String&gt;| &lt;可为 null&gt;|获取电子邮件或约会中找到的电子邮件地址。|
|`meetingSuggestions`| Array.&lt;[MeetingSuggestion](simple-types.md#meetingsuggestion)&gt;| &lt;可为 null&gt;|获取电子邮件中找到的会议建议。|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;可为 null&gt;|获取电子邮件或约会中找到的电话号码。|
|`taskSuggestions`| Array.&lt;[TaskSuggestion](simple-types.md#tasksuggestion)&gt;| &lt;可为 null&gt;|获取电子邮件或约会中找到的任务建议。|
|`urls`| Array.&lt;String&gt;| &lt;可为 null&gt;|获取电子邮件或约会中呈现的 Internet URL。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="localclienttime"></a>LocalClientTime

表示本地客户端时区中的日期和时间。仅限阅读模式。

##### <a name="properties:"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`month`| 数字|表示月份的整数值，从 0（表示 1 月）开始到 11（表示十二月）。|
|`date`| 数字|表示月份中的某一天的整数值。|
|`year`| 数字|表示年份的整数值。|
|`hours`| 数字|表示 24 小时时钟的小时的整数值。|
|`minutes`| 数字|表示分钟的整数值。|
|`seconds`| 数字|表示秒的整数值。|
|`milliseconds`| 数字|表示毫秒的整数值。|
|`timezoneOffset`| 数字|表示本地时区与 UTC 之间的分钟数差异的整数值。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="meetingsuggestion"></a>MeetingSuggestion

表示在项目中找到的建议会议。仅限阅读模式。

当对活动项目调用 [`meetingSuggestions`](simple-types.md#entities) 或 [`Entities`](Office.context.mailbox.item.md#getentities--entities) 方法时，在电子邮件中建议的会议列表将在返回的 [`getEntities`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) 对象的 `getEntitiesByType` 属性中返回。

`start` 和 `end` 的值为包含建议会议的开始和结束日期和时间的 Date 对象的字符串表示形式。这些值用为当前用户指定的默认时区表示。

##### <a name="properties:"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`attendees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|获取建议会议的与会者。|
|`end`| 字符串|获取建议会议结束的日期和时间。|
|`location`| String|获取建议会议的位置。|
|`meetingString`| String|获取标识为会议建议的字符串。|
|`start`| String|获取建议会议开始的日期和时间。|
|`subject`| String|获取建议会议的主题。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="phonenumber"></a>PhoneNumber

表示项目中标识的电话号码。仅限阅读模式。

包含电子邮件中找到的电话号码的 `PhoneNumber` 对象数组在对所选项目调用 [`phoneNumbers`](simple-types.md#entities) 方法时返回的 [`Entities`](Office.context.mailbox.item.md#getentities--entities) 对象的 `getEntities` 属性中返回。

##### <a name="type:"></a>类型：

*   对象

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`originalPhoneString`| 字符串|获取在项中识别为电话号码的文本。|
|`phoneString`| String|获取包含电话号码的字符串。该字符串只包含电话号码中的数字，而不包括原始项目中存在的括号和连字符等字符。|
|`type`| String|获取标识电话号码的类型的字符串：`Home`、`Work`、`Mobile` 和 `Unspecified`。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
#### <a name="tasksuggestion"></a>TaskSuggestion

表示项目中标识的建议任务。仅限阅读模式。

当对活动项目调用 [`taskSuggestions`](simple-types.md#entities) 或 [`Entities`](Office.context.mailbox.item.md#getentities--entities) 方法时，在电子邮件中建议的任务列表将在返回的 [`Entities`][`getEntities`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) 对象的 `getEntitiesByType` 属性中返回。

##### <a name="properties:"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`assignees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|获取应向其分配建议任务的用户。|
|`taskString`| String|获取标识为任务建议的项的文本。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 阅读|
