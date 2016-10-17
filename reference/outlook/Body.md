

# <a name="body"></a>Body

`body` 对象提供为邮件或约会添加和更新内容的方法。它在所选项的 `body` 属性中返回。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

### <a name="methods"></a>方法

####  <a name="getasync(coerciontype,-[options],-[callback])"></a>getAsync(coercionType, [options], [callback])

使用指定的格式返回当前正文。

此方法以 `coercionType` 指定的格式返回整个当前正文。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||用于返回的正文的格式。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

正文以 `asyncResult.value` 属性中请求的格式提供。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="examples"></a>示例

本示例获取纯文本格式的邮件正文。

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Do something with the result
  });
```

下面是传递给回调函数的 `result` 参数的示例。

```js
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="gettypeasync([options],-[callback])"></a>getTypeAsync([options], [callback])

获取一个值，该值指示内容采用 HTML 格式还是文本格式。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

内容类型作为 `asyncResult.value` 属性中的一个 [CoercionType](Office.md#coerciontype-string) 值返回。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|
####  <a name="prependasync(data,-[options],-[callback])"></a>prependAsync(data, [options], [callback])

将指定内容添加到项目正文开头。

`prependAsync` 方法将指定的字符串插入项目正文的开头。调用 `prependAsync` 方法的方式与调用 [`setSelectedDataAsync`](#setselecteddataasyncdata-options-callback) 方法的方式相同，插入点位于正文内容的开头。

在 HTML 标记中加入链接时，你可以通过将定位标记 (`<a>`) 上的 `id` 属性设置为 `LPNoLP` 来禁用在线链接预览。例如：

```js
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`data`| String||将插入到正文开头的字符串。字符串大小限制为 1,000,000 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;可选&gt;</td><td>主体所需的格式。<code>data</code> 参数中的字符串将转换为此格式。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>`asyncResult.error` 属性中将提供遇到的所有错误。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> 参数的长度超过 1,000,000 个字符。</td></tr></tbody></table>|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|
####  <a name="setasync(data,-[options],-[callback])"></a>setAsync(data, [options], [callback])

将整个正文替换为指定的文本。

`setAsync` 方法将项目的现有正文替换为指定的字符串，或者如果在编辑器中已选择文本，则该方法将替换选择的文本。

在 HTML 标记中加入链接时，你可以通过将定位标记 (`<a>`) 上的 `id` 属性设置为 `LPNoLP` 来禁用在线链接预览。例如：

```js
Office.context.mailbox.item.body.setAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`data`| 字符串||将替换现有正文的字符串。字符串限制为 1,000,000 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;可选&gt;</td><td>主体所需的格式。<code>data</code> 参数中的字符串将转换为此格式。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>`asyncResult.error` 属性中将提供遇到的所有错误。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> 参数的长度超过 1,000,000 个字符。</td></tr><tr><td><code>InvalidFormatError</code></td><td><code>options.coercionType</code> 参数设置为 <code>Office.CoercionType.Html</code>，而且邮件正文是纯文本格式。</td></tr></tbody></table>|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### <a name="examples"></a>示例

以下示例将正文替换为 HTML 内容。

```js
Office.context.mailbox.item.body.setAsync(
  "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
  { coercionType:"html", asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Process the result
  });
```

下面是传递给回调函数的 `result` 参数的示例。

```js
{
  "value":null,
  "status":"succeeded",
  "asyncContext":"This is passed to the callback"
}
```

####  <a name="setselecteddataasync(data,-[options],-[callback])"></a>setSelectedDataAsync(data, [options], [callback])

将正文中的所选内容更换为指定文本。

`setSelectedDataAsync` 方法将指定的字符串插入项目正文中的光标位置，或者，如果在编辑器中选定了文本，它就会替换所选文本。如果光标从未出现在项目正文中，或者如果该项目的正文不关注 UI，该字符串将插入到正文内容的顶部。

在 HTML 标记中加入链接时，你可以通过将定位标记 (`<a>`) 上的 `id` 属性设置为 `LPNoLP` 来禁用在线链接预览。例如：

```js
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`data`| String||将插入到正文中的字符串。字符串大小限制为 1,000,000 个字符。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;可选&gt;</td><td>主体所需的格式。<code>data</code> 参数中的字符串将转换为此格式。</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>`asyncResult.error` 属性中将提供遇到的所有错误。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> 参数的长度超过 1,000,000 个字符。</td></tr><tr><td><code>InvalidFormatError</code></td><td>正文类型设置为 HTML，并且数据参数包含纯文本。</td></tr></tbody></table>|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|
