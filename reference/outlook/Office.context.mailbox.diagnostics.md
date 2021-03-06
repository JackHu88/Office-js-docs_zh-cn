

# <a name="diagnostics"></a>diagnostics

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

将诊断信息提供给 Outlook 外接程序。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

### <a name="members"></a>成员

####  <a name="hostname-string"></a>hostName :String

获取表示主机应用程序的名称的字符串。

可以是下列值之一的字符串：`Outlook`、`Mac Outlook`、`OutlookIOS` 或 `OutlookWebApp`。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="hostversion-string"></a>hostVersion :String

获取表示主机应用程序或 Exchange Server 的版本的字符串。

如果邮件外接程序正在 Outlook 桌面客户端或 Outlook for iOS 上运行，则 `hostVersion` 属性返回主机应用程序版本 Outlook。在 Outlook Web App 中，属性返回 Exchange Server 的版本。其中的一个示例是字符串 `15.0.468.0`。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="owaview-string"></a>OWAView :String

获取表示 Outlook Web App 的当前视图的字符串。

返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。

如果主机应用程序不是 Outlook Web App，则访问此属性将导致返回 `undefined`。

Outlook Web App 具有三种视图，这些视图分别与屏幕和窗口的宽度以及可以显示的列数相对应：

*   `OneColumn` 在屏幕较窄时显示。Outlook Web App 在智能手机的整个屏幕上使用此单列布局。
*   `TwoColumns` 在屏幕较宽时显示。Outlook Web App 在大多数平板电脑上使用此视图。
*   `ThreeColumns` 在屏幕为宽屏时显示。例如，Outlook Web App 在台式机的全屏窗口中使用此视图。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
