# <a name="application-object-javascript-api-for-onenote"></a>Application 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_


表示包含所有全局可寻址的 OneNote 对象（如笔记本、活动笔记本和活动分区）的顶级对象。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|notebooks|[NotebookCollection](notebookcollection.md)|获取 OneNote 应用程序实例中打开的笔记本集合。在 OneNote Online 的应用程序实例中，笔记本一次仅能打开一个。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|获取活动笔记本（若有）。如果没有活动笔记本，则引发 ItemNotFound。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|获取活动笔记本（若有）。如果没有活动笔记本，则返回 NULL。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[Outline](outline.md)|获取活动边框（若有）。如果没有活动边框，则引发 ItemNotFound。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Outline](outline.md)|获取活动边框（若有）。如果没有活动边框，则返回 NULL。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[Page](page.md)|获取活动页（若有）。如果没有活动页，则引发 ItemNotFound。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|获取活动页（若有）。如果没有活动页，则返回 NULL。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[Section](section.md)|获取活动分区（若有）。如果没有活动分区，则引发 ItemNotFound。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|获取活动分区（若有）。如果没有活动分区，则返回 NULL。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page:Page)](#navigatetopagepage-page)|void|在应用程序实例中打开指定页。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|获取指定页，然后在应用程序实例中打开它。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## <a name="method-details"></a>方法详细信息


### <a name="getactivenotebook"></a>getActiveNotebook()
如果活动笔记本存在，则对其获取。如果没有处于活动状态的部分，则引发 ItemNotFound。

#### <a name="syntax"></a>语法
```js
applicationObject.getActiveNotebook();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Notebook](notebook.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivenotebookornull"></a>getActiveNotebookOrNull()
如果活动笔记本存在，则对其获取。如果没有处于活动状态的部分，则返回 null。

#### <a name="syntax"></a>语法
```js
applicationObject.getActiveNotebookOrNull();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Notebook](notebook.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactiveoutline"></a>getActiveOutline()
如果活动边框存在，则对其获取，如果没有处于活动状态的边框，则引发 ItemNotFound。

#### <a name="syntax"></a>语法
```js
applicationObject.getActiveOutline();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Outline](outline.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactiveoutlineornull"></a>getActiveOutlineOrNull()
如果活动边框存在，则对其获取，否则，返回 null。

#### <a name="syntax"></a>语法
```js
applicationObject.getActiveOutlineOrNull();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Outline](outline.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivepage"></a>getActivePage()
如果活动页面存在，则对其获取。如果没有处于活动状态的页面，则引发 ItemNotFound。

#### <a name="syntax"></a>语法
```js
applicationObject.getActivePage();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
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


### <a name="getactivepageornull"></a>getActivePageOrNull()
如果活动页面存在，则对其获取。如果没有处于活动状态的页面，则返回 null。

#### <a name="syntax"></a>语法
```js
applicationObject.getActivePageOrNull();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivesection"></a>getActiveSection()
如果活动分区存在，则对其获取。如果没有处于活动状态的分区，则引发 ItemNotFound。

#### <a name="syntax"></a>语法
```js
applicationObject.getActiveSection();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
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


### <a name="getactivesectionornull"></a>getActiveSectionOrNull()
如果活动分区存在，则对其获取。如果没有处于活动状态的分区，则返回 null。

#### <a name="syntax"></a>语法
```js
applicationObject.getActiveSectionOrNull();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void

### <a name="navigatetopagepage-page"></a>navigateToPage(page:Page)
打开应用程序实例中指定的页面。

#### <a name="syntax"></a>语法
```js
applicationObject.navigateToPage(page);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|page|Page|要打开的页面。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
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


### <a name="navigatetopagewithclienturlurl-string"></a>navigateToPageWithClientUrl(url: string)
获取特定页面，并在应用程序实例中将其打开。

#### <a name="syntax"></a>语法
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|url|string|要打开页面的客户端 url。|

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
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
