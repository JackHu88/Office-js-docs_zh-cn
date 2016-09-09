# Application 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_


表示包含所有全局可寻址的 OneNote 对象（例如笔记本、活动笔记本和活动分区）的顶级对象。

## 属性

无

## Relationships
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|笔记本|[NotebookCollection](notebookcollection.md)|获取 OneNote 应用程序实例中打开的笔记本集合。在 OneNote Online 的应用程序实例中，笔记本一次仅能打开一个。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## 方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[笔记本](notebook.md)|如果活动笔记本存在，则对其获取。 如果没有处于活动状态的部分，则引发 ItemNotFound。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[笔记本](notebook.md)|如果活动笔记本存在，则对其获取。 如果没有处于活动状态的部分，则返回 null。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[边框](outline.md)|如果活动边框存在，则对其获取，如果没有处于活动状态的边框，则引发 ItemNotFound。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[边框](outline.md)|如果活动边框存在，则对其获取，否则，返回 null。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[页面](page.md)|如果活动页面存在，则对其获取。 如果没有处于活动状态的页面，则引发 ItemNotFound。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[页面](page.md)|如果活动页面存在，则对其获取。 如果没有处于活动状态的页面，则返回 null。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[分区](section.md)|如果活动分区存在，则对其获取。 如果没有处于活动状态的分区，则引发 ItemNotFound。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[分区](section.md)|如果活动分区存在，则对其获取。 如果没有处于活动状态的分区，则返回 null。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page:Page)](#navigatetopagepage-page)|void|打开应用程序实例中指定的页面。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[页面](page.md)|获取特定页面，并在应用程序实例中将其打开。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## 方法详细信息


### getActiveNotebook()
如果活动笔记本存在，则对其获取。 如果没有处于活动状态的部分，则引发 ItemNotFound。

#### 语法
```js
applicationObject.getActiveNotebook();
```

#### 参数
无

#### 返回
[笔记本](notebook.md)

#### 示例
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


### getActiveNotebookOrNull()
如果活动笔记本存在，则对其获取。 如果没有处于活动状态的部分，则返回 null。

#### 语法
```js
applicationObject.getActiveNotebookOrNull();
```

#### 参数
无

#### 返回
[笔记本](notebook.md)

#### 示例
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


### getActiveOutline()
如果活动边框存在，则对其获取，如果没有处于活动状态的边框，则引发 ItemNotFound。

#### 语法
```js
applicationObject.getActiveOutline();
```

#### 参数
无

#### 返回
[分级显示](outline.md)

#### 示例
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


### getActiveOutlineOrNull()
如果活动边框存在，则对其获取，否则，返回 null。

#### 语法
```js
applicationObject.getActiveOutlineOrNull();
```

#### 参数
无

#### 返回
[分级显示](outline.md)

#### 示例
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


### getActivePage()
如果活动页面存在，则对其获取。 如果没有处于活动状态的页面，则引发 ItemNotFound。

#### 语法
```js
applicationObject.getActivePage();
```

#### 参数
无

#### 返回
[Page](page.md)

#### 示例
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


### getActivePageOrNull()
如果活动页面存在，则对其获取。 如果没有处于活动状态的页面，则返回 null。

#### 语法
```js
applicationObject.getActivePageOrNull();
```

#### 参数
无

#### 返回
[Page](page.md)

#### 示例
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


### getActiveSection()
如果活动分区存在，则对其获取。 如果没有处于活动状态的分区，则引发 ItemNotFound。

#### 语法
```js
applicationObject.getActiveSection();
```

#### 参数
无

#### 返回
[节](section.md)

#### 示例
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


### getActiveSectionOrNull()
如果活动分区存在，则对其获取。 如果没有处于活动状态的分区，则返回 null。

#### 语法
```js
applicationObject.getActiveSectionOrNull();
```

#### 参数
无

#### 返回
[节](section.md)

#### 示例
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


### load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

### navigateToPage(page:Page)
打开应用程序实例中指定的页面。

#### 语法
```js
applicationObject.navigateToPage(page);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|page|Page|要打开的页面。|

#### 返回
void

#### 示例
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


### navigateToPageWithClientUrl(url: string)
获取特定页面，并在应用程序实例中将其打开。

#### 语法
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|url|string|要打开页面的客户端 url。|

#### 返回
[Page](page.md)

#### 示例
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
