

# <a name="subject"></a>主题

提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|

### <a name="methods"></a>方法

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

获取约会或邮件的主题。

`getAsync` 方法开始对 Exchange 服务器进行异步调用，以获取约会或邮件的主题。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

项目主题作为 `asyncResult.value` 属性中的字符串提供。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|
####  <a name="setasync(subject,-[options],-[callback])"></a>setAsync(subject, [options], [callback])

设置约会或邮件的主题。

`setAsync` 方法开始对 Exchange 服务器进行异步调用，以设置约会或邮件的主题。设置主题将覆盖当前主题，但会保留所有前缀，如“Fwd:”或“Re:”。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`subject`| 字符串||约会或邮件的主题。字符串长度限制为 255 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>如果设置主题失败，`asyncResult.error` 属性将包含一个错误代码。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>subject</code> 参数的长度超过 255 个字符。</td></tr></tbody></table>|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|
