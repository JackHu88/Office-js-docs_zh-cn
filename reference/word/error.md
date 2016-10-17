# <a name="officeextension.error-object-(javascript-api-for-word)"></a>OfficeExtension.Error 对象（适用于 Word 的 JavaScript API）

表示使用 Word JavaScript API 时出现的错误。

_适用于：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|code|字符串|获取一个指示错误类型的值。值可以是“AccessDenied”、“GeneralException”、“ActivityLimitReached”、“InvalidArgument”、“ItemNotFound”或“NotImplemented”。 <!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|获取指示出错时所发生的问题的一个值。此值仅在开发/调试过程中使用。  |
|邮件 |字符串| 获取与错误代码对应的本地化的人工读取字符串。|
|name |字符串| 获取一个始终为“OfficeExtension.Error”的值。 |
|traceMessages |string[]| 获取值数组，这些值与通过 context.trace(); 设置的检测消息对应 |

_请参阅属性访问[示例](#property-access-examples)_。

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|以下面的格式返回错误代码和消息值：“{0}: {1}”、代码、消息。|

## <a name="method-details"></a>方法详细信息

### <a name="tostring()"></a>toString()
以下面的格式返回错误代码和消息值：“{0}: {1}”、代码、消息。

#### <a name="syntax"></a>语法
```js
error.toString()
```

#### <a name="parameters"></a>参数
无。

#### <a name="returns"></a>返回
string

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
    body.insertText(0);

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Error code and message: ' + error.toString());
    }
});

```

## <a name="property-access-examples"></a>属性访问示例

### <a name="trace-message-instrumentation"></a>跟踪消息检测

下面的示例显示如何检测批处理命令，以确定错误发生的位置。第一批成功插入文档中的前两个段落，未导致任何错误。第二批成功插入第三和第四段落，但在调用以插入第五段时失败。批处理中在该失败命令之后的所有其他命令都不执行，包括添加第五个跟踪消息的命令。在这种情况下，插入第四段之后及添加第五个跟踪消息之前，出现了错误。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
    body.insertParagraph('1st paragraph', Word.InsertLocation.end);
    // Queue a command for instrumenting this part of the batch.
    context.trace('1st paragraph successful');

    body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
    context.trace('2nd paragraph successful');

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
        body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
        context.trace('3rd paragraph successful');

        body.insertParagraph('4th paragraph', Word.InsertLocation.end);
        context.trace('4th paragraph successful');

        // This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
        body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
        // Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
        context.trace('5th paragraph successful');
    }).then(context.sync);
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Trace messages: ' + error.traceMessages);
    }
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
