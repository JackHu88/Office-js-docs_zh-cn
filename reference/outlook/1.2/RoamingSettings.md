

# <a name="roamingsettings"></a>RoamingSettings

通过使用 `RoamingSettings` 对象的方法创建的设置将按外接程序和按用户进行保存。即，这些设置仅供创建它们的外接程序使用，并且仅来自保存它们的用户邮箱。

> 虽然 Outlook 外接程序 API 仅允许创建它们的外接程序访问这些设置，但这些设置不应被视为安全存储。可以通过 Exchange Web 服务或扩展 MAPI 访问这些设置。它们不应用于存储敏感信息，如用户凭据或安全令牌。

设置的名称是一个字符串，而值可以是字符串、数字、布尔值、null 值、对象或数组。

可通过 `Office.context` 命名空间中的 [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) 属性访问 `RoamingSettings` 对象。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|

### <a name="example"></a>示例

```JavaScript
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### <a name="methods"></a>方法

####  <a name="get(name)-→-(nullable)-{string|number|boolean|object|array}"></a>get(name) → (nullable) {String|Number|Boolean|Object|Array}

检索指定设置。

##### <a name="parameters:"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|要检索的设置的区分大小写的名称。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="returns:"></a>返回：

<dl class="param-type">

<dt>
类型</dt>


<dd>字符串 | 数字 | 布尔值 | 对象 | 数组</dd>

</dl>

####  <a name="remove(name)"></a>remove(name)

移除指定设置。

##### <a name="parameters:"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|要移除的设置的区分大小写的名称。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="saveasync([callback])"></a>saveAsync([callback])

保存设置。

外接程序初始化时会加载之前保存的所有设置，因此，在会话的生存期内，只能通过 [`set`](RoamingSettings.md#set) 和 [`get`](RoamingSettings.md#get) 方法使用设置属性包的内存副本。如果希望保留这些设置以便可在下次使用外接程序时使用，请使用 `saveAsync` 方法。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 |

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="set(name,-value)"></a>set(name, value)

设置或创建指定设置。

set 方法创建指定名称的新设置（如果该设置尚不存在），或者设置指定名称的现有设置。该值在文档中存储为其数据类型的序列化 JSON 表示形式。

每个外接程序的设置的最大可用空间为 2 MB，并且每个单独的设置的空间限制为 32 KB。

在调用 [`saveAsync`](RoamingSettings.md#saveasynccallback) 函数之前，使用 `set` 函数对设置所做的所有更改将不会保存到服务器。

##### <a name="parameters:"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|要设置或创建的设置的名称（区分大小写）。|
|`value`| 字符串 &#124; 数字 &#124; 布尔值 &#124; 对象 &#124; 数组|要存储的值。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|