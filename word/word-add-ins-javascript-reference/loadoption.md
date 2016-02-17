# LoadOption 对象（适用于 Word 的 JavaScript API）

指定调用 context.sync() 时要加载的分页信息和属性的对象。 

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
| 属性   | 类型|说明|
|:---------------|:--------|:----------|
|select|object|包含参数/关系名称的逗号分隔列表或数组。可选。|
|expand|object|包含关系名称的逗号分隔列表或数组。可选。|
|top|int| 指定结果中可以包含的集合项最大数量。可选。|
|skip|int|指定要跳过且不包含在结果中的集合中的项数目。如果指定 `top`，跳过指定数目的项目后将会启动结果集。可选。|

## 详细信息

指定属性和分页信息的首选方法时使用字符串文本。前两个示例说明了请求段落集合中段落的文本和字体大小属性的首选方法：

<code>context.load(paragraphs, 'text, font/size, top:50, skip:0');</code>

<code>paragraphs.load('text, font/size, top:50, skip:0');</code>

下面是使用对象表示法的等效方法：

&lt;code&gt;context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>
                                
&lt;code&gt;paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

请注意，如果我们未在 select 语句中指定字体对象的特定属性，expand 语句本身将指示需加载所有字体属性。 

## 示例

本示例说明如何获取 Word 文档中的前 50 个段落及其文本和字体大小属性。

```js
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the text and font properties for the top 50 paragraphs.
            // It is best practice to always specify the property set. Otherwise, all properties are
            // returned in on the object. 
            context.load(paragraphs, 'text, font/size, top: 50, skip: 0');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
            
            // Insert code that works with the paragraphs loaded by context.load().

        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

```

## 支持详细信息

在运行时检查过程中使用[要求设置](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)。 
