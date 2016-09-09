

# Location

提供用于获取和设置 Outlook 外接程序中的会议地点的方法。

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|

### 方法

####  getAsync([options], callback)

获取约会的位置。

`getAsync` 方法开始对 Exchange 服务器进行异步调用，以获取约会的位置。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

约会的位置以 `asyncResult.value` 属性中的字符串形式提供。|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|
####  setAsync(location, [options], [callback])

设置约会的位置。

`setAsync` 方法开始对 Exchange 服务器进行异步调用，以设置约会的位置。设置约会的位置将覆盖当前位置。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`location`| String||约会的位置。字符串大小限制为 255 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>如果设置位置失败，`asyncResult.error` 属性将包含一个错误代码。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>location</code> 参数的长度超过 255 个字符。</td></tr></tbody></table>|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|
