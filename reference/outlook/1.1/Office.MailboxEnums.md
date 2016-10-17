 

# <a name="mailboxenums"></a>MailboxEnums

## [Office](Office.md).MailboxEnums

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写|

### <a name="members"></a>成员

#### <a name="attachmenttype-:string"></a>AttachmentType :String

指定附件的类型。仅限撰写模式。

AttachmentType

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`File`| String|附件是一个文件。|
|`Item`| 字符串|附件是一个 Exchange 项目。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写|
#### <a name="entitytype-:string"></a>EntityType :String

指定实体的类型。仅限撰写模式。

EntityType

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`Address`| 字符串|指定实体为通讯地址。|
|`Contact`| String|指定实体为联系人。|
|`EmailAddress`| String|指定实体为 SMTP 电子邮件地址。|
|`MeetingSuggestion`| String|指定实体为会议建议。|
|`PhoneNumber`| String|指定实体为美国电话号码。|
|`TaskSuggestion`| String|指定实体为任务建议。|
|`URL`| String|指定实体为 Internet URL。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写|
#### <a name="itemtype-:string"></a>ItemType :String

指定项目的类型。仅限撰写模式。

ItemType

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`Message`| String|电子邮件、会议请求、会议响应或会议取消。|
|`Appoinment`| String|约会项目。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写|
#### <a name="recipienttype-:string"></a>RecipientType :String

指定约会收件人的类型。仅限撰写模式。

RecipientType

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`Other`| String|收件人不是其他收件人类型之一。|
|`DistributionList`| String|收件人是包含电子邮件地址列表的通讯组列表。|
|`User`| String|收件人是位于 Exchange 服务器上的 SMTP 电子邮件地址。|
|`ExternalUser`| String|收件人是不位于 Exchange 服务器上的 SMTP 电子邮件地址。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|适用的 Outlook 模式| 撰写|
#### <a name="responsetype-:string"></a>ResponseType :String

指定对会议邀请的响应类型。仅限撰写模式。

ResponseType

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`None`| String|参与者尚未响应。|
|`Organizer`| String|参与者是会议组织者。|
|`Tentative`| 字符串|参与者暂时接受会议请求。|
|`Accepted`| String|参与者已接受会议请求。|
|`Declined`| 字符串|参与者已拒绝会议请求。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写|
