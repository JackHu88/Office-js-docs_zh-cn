

# 时间

`Time` 对象作为撰写模式中约会的 [`start`](Office.context.mailbox.item.md#start-datetime) 或 [`end`](Office.context.mailbox.item.md#end-datetime) 属性返回。

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|

### 方法

####  getAsync([options], callback)

获取约会的开始或结束时间。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。

日期和时间作为 `asyncResult.value` 属性中的 Date 对象提供。该值以协调世界时 (UTC) 表示。您可以使用 [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法|将 UTC 时间转换为本地客户端时间。

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写|
####  setAsync(dateTime, [options], [callback])

设置约会的开始或结束时间。

如果对 [`setAsync`](Office.context.mailbox.item.md#start-datetime) 属性调用了 `start` 方法，[`end`](Office.context.mailbox.item.md#end-datetime) 属性将会调整为维持约会持续时间（同之前的设置一样）。如果对 `setAsync` 属性调用了`end` 方法，则约会的持续时间将延长到新的结束时间。

时间必须以 UTC 格式表示；您可以使用 [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) 方法获取正确的 UTC 时间。

##### 参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`dateTime`| 日期||Date 对象以协调世界时 (UTC) 表示。|
|`options`| Object| &lt;可选&gt;|包含一个或多个以下属性的对象文字。<br/><br/>**属性**<br/><table class="nested-table"><thead><tr><th>名称</th><th>类型</th><th>属性</th><th>说明</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;可选&gt;</td><td>开发人员可以提供他们想要在回调方法中访问的任何对象。</td></tr></tbody></table>|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](simple-types.md#asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。 <br/>如果设置日期和时间失败，`asyncResult.error` 属性将包含一个错误代码。<br/><table class="nested-table"><thead><tr><th>错误代码</th><th>说明</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>约会结束时间早于约会开始时间。</td></tr></tbody></table>|

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|适用的 Outlook 模式| 撰写|

##### 示例

下面的示例设置约会的开始时间。

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```
