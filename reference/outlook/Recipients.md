

# <a name="recipients"></a>Recipients

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|

### <a name="methods"></a>方法

####  <a name="addasync(recipients,-[options],-[callback])"></a>addAsync(recipients, [options], [callback])

将收件人列表添加到约会或邮件的现有收件人中。

`recipients` 参数可以是以下任何一个数组：

*   包含 SMTP 电子邮件地址的字符串
*   `EmailUser` 对象
*   `EmailAddressDetails` 对象

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||要添加到收件人列表中的收件人。|
|`options`| 对象| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果添加收件人失败，`asyncResult.error` 属性将包含一个错误代码。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>收件人的数量超过 100 个条目。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

下面的示例创建 `EmailUser` 对象的数组，并将其添加到邮件收件人的“收件人”中。

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.to.addAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients added");
  }
});
```

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

获取约会或邮件的收件人列表。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。

在调用完成后，`asyncResult.value` 属性将包含 [`EmailAddressDetails`](simple-types.md#emailaddressdetails) 对象的数组。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

下面的示例获取会议的可选与会者。

```js
Office.context.mailbox.item.optionalAttendees.getAsync(function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    var msg = "";
    result.value.forEach(function(recip, index) {
      msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
    });
    showMessage(msg);
  }
});
```

####  <a name="setasync(recipients,-[options],-callback)"></a>setAsync(recipients, [options], callback)

设置约会或邮件的收件人列表。

`setAsync` 方法将覆盖当前收件人列表。

`recipients` 参数可以是以下任何一个数组：

*   包含 SMTP 电子邮件地址的字符串
*   `EmailUser` 对象
*   `EmailAddressDetails` 对象

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||要添加到收件人列表中的收件人。|
|`options`| 对象| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。 <br/>如果设置收件人失败，`asyncResult.error` 属性将包含一个代码，表示在添加数据时出现的所有错误。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>收件人的数量超过 100 个条目。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="example"></a>示例

下面的示例创建 `EmailUser` 对象的数组，并用该数组替换邮件的“抄送”收件人。

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.cc.setAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients overwritten");
  }
});
```
