# RequestContext 对象（适用于 Word 的 JavaScript API）

RequestContext 可加快从 Word 外接程序发出的对 Word 应用程序的请求，因为这两个应用程序在不同的进程中运行。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
无

## 方法

| 方法         | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |使用参数指定的属性和选项填充在 JavaScript 层中创建的代理对象。|
|[sync()](#sync)  |Promise 对象 |将请求队列提交到 Word 并返回一个 promise 对象，此对象可用于将其他操作链接起来。|

## 方法详细信息

### load(object: object, option: object)
使用参数指定的属性和选项填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
requestContextObject.load(object, loadOption);
```

#### 参数
| 参数       | 类型    |说明|
|:----------------|:--------|:----------|
|object|object|可选。指定要加载的对象的名称。|
|选项|[loadOption](loadoption.md)|可选，但可作为最佳实践。指定加载选项，例如选择、展开、跳过和置顶。 |

#### 返回
void

##### 示例

下面的示例说明如何将请求上下文用于加载段落集合上的文本属性。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

#### 其他信息

添加跟踪的对象后，必须调用 load()。

### sync()
将请求队列提交到 Word 并返回一个 promise 对象，此对象可用于将其他操作链接起来。

#### 语法
```js
requestContextObject.sync();
```

#### 参数
无

#### 返回
Promise 对象。

#### 示例

下面的示例显示使用了两次的同步方法：1) 加载内容控件集合，其中包含每个内容控件的文本属性，2) 清除集合中第一个内容控件的内容。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the content controls collection.
    contentControls.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {

            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });
        }

    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

## 支持详细信息
在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。