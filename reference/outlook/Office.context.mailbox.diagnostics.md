

# diagnostics

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

将诊断信息提供给 Outlook 外接程序。

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

### 成员

####  hostName :String

获取表示主机应用程序的名称的字符串。

可以是下列值之一的字符串：`Outlook`、`Mac Outlook` 或 `OutlookWebApp`。

##### 类型：

*   String

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  hostVersion :String

获取表示主机应用程序或 Exchange Server 的版本的字符串。

如果邮件外接程序运行在 Outlook 桌面客户端上，则 `hostVersion` 属性返回主机应用程序 Outlook 的版本。在 Outlook Web App 中，该属性返回 Exchange Server 的版本。例如，字符串 `15.0.468.0`。

##### 类型：

*   String

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  OWAView :String

获取表示 Outlook Web App 的当前视图的字符串。

返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。

如果主机应用程序不是 Outlook Web App，则访问此属性将导致返回 `undefined`。

Outlook Web App 具有三种视图，这些视图分别与屏幕和窗口的宽度以及可以显示的列数相对应：

*   `OneColumn` 在屏幕较窄时显示。Outlook Web App 在智能手机的整个屏幕上使用此单列布局。
*   `TwoColumns` 在屏幕较宽时显示。Outlook Web App 在大多数平板电脑上使用此视图。
*   `ThreeColumns` 在屏幕为宽屏时显示。例如，Outlook Web App 在台式机的全屏窗口中使用此视图。

##### 类型：

*   String

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
