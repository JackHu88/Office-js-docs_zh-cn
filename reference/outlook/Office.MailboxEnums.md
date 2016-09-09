 

# MailboxEnums

## [Office](Office.md).MailboxEnums

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|

### 成员

#### AttachmentType :String

指定附件的类型。

AttachmentType

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`File`| String|`file`|附件是一个文件。|
|`Item`| 字符串|`item`|附件是一个 Exchange 项目。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
#### EntityType :String

指定实体的类型。

EntityType

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`Address`| 字符串|`address`|指定实体为通讯地址。|
|`Contact`| String|`contact`|指定实体为联系人。|
|`EmailAddress`| String|`emailAddress`|指定实体为 SMTP 电子邮件地址。|
|`MeetingSuggestion`| String|`meetingSuggestion`|指定实体为会议建议。|
|`PhoneNumber`| String|`phoneNumber`|指定实体为美国电话号码。|
|`TaskSuggestion`| String|`taskSuggestion`|指定实体为任务建议。|
|`URL`| String|`url`|指定实体为 Internet URL。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
#### ItemNotificationMessageType :String

为约会或邮件指定通知邮件类型。

ItemNotificationMessageType

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|notificationMessage 是进度指示器。|
|`InformationalMessage`| 字符串|`informationalMessage`|notificationMessage 是信息性消息。|
|`ErrorMessage`| 字符串|`errorMessage`|notificationMessage 是错误消息。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|适用的 Outlook 模式| 撰写或阅读|
#### ItemType :String

指定项的类型。

ItemType

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`Message`| String|`message`|电子邮件、会议请求、会议响应或会议取消。|
|`Appointment`| String|`appointment`|约会项目。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
#### RecipientType :String

指定约会收件人的类型。

RecipientType

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`Other`| String|`other`|收件人不是其他收件人类型之一。|
|`DistributionList`| String|`distributionList`|收件人是包含电子邮件地址列表的通讯组列表。|
|`User`| String|`user`|收件人是位于 Exchange 服务器上的 SMTP 电子邮件地址。|
|`ExternalUser`| String|`externalUser`|收件人是不位于 Exchange 服务器上的 SMTP 电子邮件地址。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|适用的 Outlook 模式| 撰写或阅读|
#### ResponseType :String

指定对会议邀请的响应类型。

ResponseType

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`None`| String|`none`|参与者尚未响应。|
|`Organizer`| String|`organizer`|参与者是会议组织者。|
|`Tentative`| 字符串|`tentative`|参与者暂时接受会议请求。|
|`Accepted`| String|`accepted`|参与者已接受会议请求。|
|`Declined`| 字符串|`declined`|参与者已拒绝会议请求。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|

#### RestVersion :String

指定对应于 REST 格式的项目 ID 的 REST API 的版本。 

RestVersion

##### 类型：

*   字符串

##### 属性：

|名称| 类型| 值 | 描述|
|---|---|---|---|
|`v1_0`| 字符串|`v1.0`|版本 1.0|
|`v2_0`| 字符串|`v2.0`|版本 2.0|
|`Beta`| 字符串|`beta`|Beta.|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|适用的 Outlook 模式| 撰写或阅读|
