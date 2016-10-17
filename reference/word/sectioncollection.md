# <a name="sectioncollection-object-(javascript-api-for-word)"></a>SectionCollection 对象（适用于 Word 的 JavaScript API）

包含文档的 [section](section.md) 对象的集合。

_适用于：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|项目|[Section[]](section.md)|分区对象的集合。只读。|

## <a name="relationships"></a>Relationships
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息

### <a name="load(param:-object)"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;

    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Create a proxy object the primary header of the first section.
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");

        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);

        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
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

## <a name="support-details"></a>支持详细信息
在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。
