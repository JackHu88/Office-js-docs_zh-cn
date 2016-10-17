# <a name="sectiongroup-object-(javascript-api-for-onenote)"></a>SectionGroup 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_   


表示 OneNote 分区组。分区组可包含分区和其他分区组。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|clientUrl{|字符串|分区组的客户端 url。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-clientUrl{)|
|id|字符串|获取分区组的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-id)|
|name|字符串|获取分区组的名称。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-name)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|获取包含分区组的笔记本。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-notebook)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|获取包含分区组的分区组。如果分区组是笔记本的直接子级，则引发 ItemNotFound。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|获取包含分区组的分区组。如果分区组是笔记本的直接子级，则返回 null。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroupOrNull)|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|分区组中的分区组集合。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|分区组中的分区集合。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sections)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[addSection(title:String)](#addsectiontitle-string)|[Section](section.md)|将新分区添加至分区组结尾。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSection)|
|[addSectionGroup(name:String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|将新的分区组添加至此 sectionGroup 结尾。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSectionGroup)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-load)|

## <a name="method-details"></a>方法详细信息


### <a name="addsection(title:-string)"></a>addSection(title:String)
将新分区添加至分区组结尾。

#### <a name="syntax"></a>语法
```js
sectionGroupObject.addSection(title);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|title|字符串|新节的名称。|

#### <a name="returns"></a>返回
[Section](section.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;
    
    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Add a section to each section group.
            $.each(sectionGroups.items, function(index, sectionGroup) {
                sectionGroup.addSection("Agenda");
            });
            
            // Run the queued commands.
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


### <a name="addsectiongroup(name:-string)"></a>addSectionGroup(name:String)
将新的分区组添加至此 sectionGroup 结尾。

#### <a name="syntax"></a>语法
```js
sectionGroupObject.addSectionGroup(name);
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
    var sectionGroup;
    var nestedSectionGroup;

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroups = notebook.sectionGroups;

    // Queue a command to load the new section group.
    sectionGroups.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function(){
            sectionGroup = sectionGroups.items[0];
            sectionGroup.load();
            return context.sync();
        })
        .then(function(){
            nestedSectionGroup = sectionGroup.addSectionGroup("Sample nested section group");
            nestedSectionGroup.load();
            return context.sync();
        })
        .then(function() {
            console.log("New nested section group name is " + nestedSectionGroup.name);
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
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group. 
    // For best performance, request specific properties.           
    sectionGroup.load("id,name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Section group ID: " + sectionGroup.id);
            
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
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group with the specified properties.           
    sectionGroup.load("name,notebook/name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Parent notebook name: " + sectionGroup.notebook.name);
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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("name");
    
    // Get the child section groups of the first section group in the notebook.
    var nestedSectionGroups = sectionGroups._GetItem(0).sectionGroups;
    
    // Queue a command to load the ID and name properties of the child section groups.
    nestedSectionGroups.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each child section group.
            $.each(nestedSectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
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

    // Get the sections that are siblings of the current section.
    var sections = context.application.getActiveSection().parentSectionGroup.sections;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sections.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each section.
            $.each(sections.items, function(index, section) {
                console.log("Section name: " + section.name);  
                console.log("Section ID: " + section.id);  
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

