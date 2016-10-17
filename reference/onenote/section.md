# <a name="section-object-(javascript-api-for-onenote)"></a>分区对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_   


表示 OneNote 分区。分区可包含页面。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|clientUrl|字符串|分区的客户端 url。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-clientUrl)|
|id|字符串|获取分区的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-id)|
|name|字符串|获取分区的名称。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-name)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|获取包含分区的笔记本。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-notebook)|
|pages|[PageCollection](pagecollection.md)|分区中的页面集合。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-pages)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|获取包含分区的分区组。如果分区是笔记本的直接子级，则引发 ItemNotFound。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|获取包含分区的分区组。如果分区是笔记本的直接子级，则返回 null。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroupOrNull)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[addPage(title: string)](#addpagetitle-string)|[Page](page.md)|添加新页面至分区结尾。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-addPage)|
|[copyToNotebook(destinationNotebook:Notebook)](#copytonotebookdestinationnotebook-notebook)|[Section](section.md)|将此分区复制到指定的笔记本。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToNotebook)|
|[copyToSectionGroup(destinationSectionGroup: SectionGroup)](#copytosectiongroupdestinationsectiongroup-sectiongroup)|[Section](section.md)|将此分区复制到指定的分区组。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToSectionGroup)|
|[insertSectionAsSibling(location: string, title: string)](#insertsectionassiblinglocation-string-title-string)|[Section](section.md)|在当前分区之前或之后插入一个新的分区。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-insertSectionAsSibling)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-load)|

## <a name="method-details"></a>方法详细信息


### <a name="addpage(title:-string)"></a>addPage(title: string)
添加新页面至分区结尾。

#### <a name="syntax"></a>语法
```js
sectionObject.addPage(title);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|title|字符串|新页面的标题。|

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {
            
    // Queue a command to add a page to the current section.
    var page = context.application.getActiveSection().addPage("Wish list");
            
    // Queue a command to load the id and title of the new page. 
    // This example loads the new page so it can read its properties later.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Page name: " + page.title);
            console.log("Page ID: " + page.id);

        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="copytonotebook(destinationnotebook:-notebook)"></a>copyToNotebook(destinationNotebook:Notebook)
将此分区复制到指定的笔记本。

#### <a name="syntax"></a>语法
```js
sectionObject.copyToNotebook(destinationNotebook);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|destinationNotebook|Notebook|要将此分区复制到的笔记本。|

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {
    var app = context.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return context.sync()
        .then(function() {
            newSection = section.copyToNotebook(notebook);
            newSection.load('id');
            return context.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="copytosectiongroup(destinationsectiongroup:-sectiongroup)"></a>copyToSectionGroup(destinationSectionGroup: SectionGroup)
将此分区复制到指定的分区组。

#### <a name="syntax"></a>语法
```js
sectionObject.copyToSectionGroup(destinationSectionGroup);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|destinationSectionGroup|分区组|要将此分区复制到的分区组。|

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (ctx) {
    var app = ctx.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return ctx.sync()
        .then(function() {
            var firstSectionGroup = notebook.sectionGroups.items[0];
            newSection = section.copyToSectionGroup(firstSectionGroup);
            newSection.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertsectionassibling(location:-string,-title:-string)"></a>insertSectionAsSibling(location: string, title: string)
在当前分区之前或之后插入一个新的分区。

#### <a name="syntax"></a>语法
```js
sectionObject.insertSectionAsSibling(location, title);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|location|string|相对于当前分区的新分区的位置。可能的值是：Before、After|
|职位|字符串|新节的名称。|

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {
            
    // Queue a command to insert a section after the current section.
    var section = context.application.getActiveSection().insertSectionAsSibling("After", "New section");
            
    // Queue a command to load the id and name of the new section. 
    // This example loads the new section so it can read its properties later.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


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

**id**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load("id");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section ID: " + section.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name 和 notebook**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section with the specified properties. 
    section.load("name,notebook/name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section name: " + section.name);
            console.log("Parent notebook name: " + section.notebook.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentSectionGroupOrNull**
```js
OneNote.run(function (context) {
    // Queue a command to add a page to the current section.
    var section = context.application.getActiveSection();
    section.load('clientUrl,notebook');
    var sectionGroup = section.parentSectionGroupOrNull;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(sectionGroup.isNull === false)
            {
                // If a parent section group exists, queue a command to add a section in it!
                sectionGroup.addSection("NewSectionInSectionGroup");
            }
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
    
