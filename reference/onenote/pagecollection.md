# <a name="pagecollection-object-(javascript-api-for-onenote)"></a>PageCollection 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示页面的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|count|int|返回集合中页面的数目。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-count)|
|items|[Page[]](page.md)|页面对象的集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-items)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|获取具有指定标题的页面集合。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getByTitle)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Page](page.md)|按其在集合中的 ID 或索引获取页面。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Page](page.md)|按其在集合中的位置获取页面。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getbytitle(title:-string)"></a>getByTitle(title: string)
获取具有指定标题的页面集合。

#### <a name="syntax"></a>语法
```js
pageCollectionObject.getByTitle(title);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|title|字符串|页面的标题。|

#### <a name="returns"></a>返回
[PageCollection](pagecollection.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get all the pages in the current section.
    var allPages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.
    allPages.load("id"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Get the sections with the specified name.
            var todoPages = allPages.getByTitle("Todo list");

            // Queue a command to load the section. 
            // For best performance, request specific properties.
            todoPages.load("id,title"); 

            return context.sync()
                .then(function () {

                    // Iterate through the collection or access items individually by index.
                    if (todoPages.items.length > 0) {
                        console.log("Page title: " + todoPages.items[0].title);
                        console.log("Page ID: " + todoPages.items[0].id);
                    }
                });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
按其在集合中的 ID 或索引获取页面。只读。

#### <a name="syntax"></a>语法
```js
pageCollectionObject.getItem(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number or string|页面的 ID 或集合中页面的索引位置。|

#### <a name="returns"></a>返回
[Page](page.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
按其在集合中的位置获取页面。

#### <a name="syntax"></a>语法
```js
pageCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[Page](page.md)

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
### <a name="property-access-examples"></a>属性访问示例

**items**
```js
OneNote.run(function (context) {
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
    
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.
            $.each(pages.items, function(index, page) {
                console.log(page.title);
                console.log(page.id);
            });
        }); 
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

