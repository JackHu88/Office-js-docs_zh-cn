# <a name="notebook-object-(javascript-api-for-onenote)"></a>笔记本对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_   


表示一个 OneNote 笔记本。笔记本包含分区组合和分区。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|clientUrl|字符串|笔记本的客户端 url只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-clientUrl)|
|id|字符串|获取笔记本的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-id)|
|name|字符串|获取笔记本的名称。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-name)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|笔记本中的分区组。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|笔记本中的分区。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sections)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[addSection(name:String)](#addsectionname-string)|[Section](section.md)|添加新分区至笔记本结尾。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSection)|
|[addSectionGroup(name:String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|将新的分区组添加到笔记本结尾。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSectionGroup)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-load)|

## <a name="method-details"></a>方法详细信息


### <a name="addsection(name:-string)"></a>addSection(name:String)
添加新分区至笔记本结尾。

#### <a name="syntax"></a>语法
```js
notebookObject.addSection(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|name|字符串|新节的名称。|

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section. 
    var section = notebook.addSection("Sample section");
    
    // Queue a command to load the new section. This example reads the name property later.
    section.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section name is " + section.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### <a name="addsectiongroup(name:-string)"></a>addSectionGroup(name:String)
将新的分区组添加到笔记本结尾。

#### <a name="syntax"></a>语法
```js
notebookObject.addSectionGroup(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|name|字符串|新节的名称。|

#### <a name="returns"></a>返回
[SectionGroup](sectiongroup.md)

#### <a name="examples"></a>示例
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroup = notebook.addSectionGroup("Sample section group");

    // Queue a command to load the new section group.
    sectionGroup.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section group name is " + sectionGroup.name);
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
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
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

**name**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook name: " + notebook.name);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**sectionGroups**
```js          
OneNote.run(function (context) {

    // Get the section groups in the notebook. 
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the sectionGroups. 
    sectionGroups.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);
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

**sections**
```js
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();
    
    // Queue a command to get immediate child sections of the notebook. 
    var childSections = notebook.sections;

    // Queue a command to load the childSections. 
    context.load(childSections);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(childSections.items, function(index, childSection) {
                console.log("Immediate child section name: " + childSection.name);
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

