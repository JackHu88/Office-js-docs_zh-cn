

# <a name="customproperties"></a>CustomProperties

`CustomProperties` 对象表示特定于某个特定项目和特定于 Outlook 的某个邮件外接程序的自定义属性。例如，邮件外接程序可能有必要保存一些特定于激活外接程序的当前电子邮件的数据。如果用户以后再次访问相同的邮件，并再次激活邮件外接程序，外接程序将能够检索作为自定义属性保存的数据。

由于 Outlook for Mac 不缓存自定义属性，因此如果用户的网络出现故障，邮件外接程序将无法访问它们的自定义属性。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

### <a name="example"></a>示例

以下示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法将这些属性重新保存到服务器。加载自定义属性后，该示例将使用 [`get`](CustomProperties.md#get) 方法读取自定义属性 `myProp`，使用 [`set`](CustomProperties.md#set) 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。

```
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

### <a name="methods"></a>方法

####  <a name="get(name)-→-{string}"></a>get(name) → {String}

返回指定自定义属性的值。

##### <a name="parameters:"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|要返回的自定义属性的名称。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="returns:"></a>返回：

指定的自定义属性的值。

<dl class="param-type">

<dt>
类型</dt>


<dd>字符串</dd>

</dl>

####  <a name="remove(name)"></a>remove(name)

从自定义属性集合中移除指定的属性。

若要永久移除属性，必须调用 `CustomProperties` 对象的 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法。

##### <a name="parameters:"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|要移除的属性的名称。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|
####  <a name="saveasync([callback],-[asynccontext])"></a>saveAsync([callback], [asyncContext])

将特定于项目的自定义属性保存到服务器。

必须调用 `saveAsync` 方法来保留使用 `CustomProperties` 对象的 [`set`](CustomProperties.md#set) 方法或 [`remove`](CustomProperties.md#remove) 方法所做的任何更改。保存操作是异步操作。

最好让你的回调函数检查并处理 `saveAsync` 中的错误。尤其要注意的是，当用户在阅读窗体中处于连接状态时，可以激活阅读外接程序，随后用户将断开连接。如果外接程序在断开状态下调用 `saveAsync`，`saveAsync` 将返回错误。你的回调方法应该会相应地处理此错误。

##### <a name="parameters:"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| 函数| &lt;可选&gt;|方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](simple-types.md#asyncresult) 对象）调用在 `callback` 参数中传递的函数。 |
|`asyncContext`| 对象| &lt;可选&gt;|传递给回调方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|

##### <a name="example"></a>示例

以下 JavaScript 代码示例显示如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性，以及如何使用 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法将这些属性保存回服务器中。加载自定义属性后，该代码示例将使用 [`get`](CustomProperties.md#get) 方法读取自定义属性 `myProp`，使用 [`set`](CustomProperties.md#set) 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="set(name,-value)"></a>set(name, value)

将指定属性设置为指定值。

`set` 方法将指定属性设置为指定值。必须使用 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法将该属性保存到服务器。

如果尚不存在指定属性，`set` 方法将创建一个新的属性；否则现有值将替换为新值。`value` 参数可以是任何类型；但是始终作为字符串传递给服务器。

##### <a name="parameters:"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`name`| 字符串|要设置的属性的名称。|
|`value`| 对象|要设置的属性的值。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../tutorial-api-requirement-sets.md)| 1.0|
|[最低权限级别](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|适用的 Outlook 模式| 撰写或阅读|