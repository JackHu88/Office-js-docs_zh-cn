 

# <a name="office"></a>Office

该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](../shared/shared-api.md)。

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|

### <a name="namespaces"></a>命名空间

[context](Office.context.md)：提供 Office 外接程序 API 的上下文命名空间中的共享接口以便在 Outlook 外接程序 API 中使用。

[MailboxEnums](Office.MailboxEnums.md)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。

### <a name="members"></a>成员

####  <a name="asyncresultstatus-:string"></a>AsyncResultStatus :String

指定异步调用的结果。

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`Succeeded`| String|调用成功。|
|`Failed`| 字符串|调用失败。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="coerciontype-:string"></a>CoercionType :String

指定如何强制由调用方法返回或设置的数据。

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`Html`| String|请求以 HTML 格式返回的数据。|
|`Text`| 字符串|请求以文本格式返回的数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="sourceproperty-:string"></a>SourceProperty :String

指定由调用方法返回的数据源。

##### <a name="type:"></a>类型：

*   字符串

##### <a name="properties:"></a>属性：

|名称| 类型| 描述|
|---|---|---|
|`Body`| 字符串|数据源来自邮件的正文。|
|`Subject`| String|数据源来自邮件的主题。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|适用的 Outlook 模式| 撰写或阅读|
