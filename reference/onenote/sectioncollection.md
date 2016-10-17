# <a name="sectioncollection-object-(javascript-api-for-onenote)"></a>SectionCollection 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示分区的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|count|int|返回集合中的分区数。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-count)|
|items|[Section[]](section.md)|分区对象的集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-items)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|获取具有指定名称的分区的集合。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Section](section.md)|按其在集合中的 ID 或索引获取分区。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Section](section.md)|按其在集合中的位置获取分区。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getbyname(name:-string)"></a>getByName(name: string)
获取具有指定名称的分区的集合。

#### <a name="syntax"></a>语法
```js
sectionCollectionObject.getByName(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|name|字符串|节的名称。|

#### <a name="returns"></a>返回
[SectionCollection](sectioncollection.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("id"); 
    
    // Get the sections with the specified name.
    var groceriesSections = sections.getByName("Groceries");
    
    // Queue a command to load the sections with the specified name.
    groceriesSections.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (groceriesSections.items.length > 0) {
                console.log("Section name: " + groceriesSections.items[0].name);
                console.log("Section ID: " + groceriesSections.items[0].id);
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

### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
按其在集合中的 ID 或索引获取分区。只读。

#### <a name="syntax"></a>语法
```js
sectionCollectionObject.getItem(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number or string|分区组的 ID 或在集合中的分区的索引位置。|

#### <a name="returns"></a>返回
[Section](section.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
按其在集合中的位置获取分区。

#### <a name="syntax"></a>语法
```js
sectionCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[Section](section.md)

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

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sections.items[0]
            $.each(sections.items, function(index, section) {
                if (section.name === "Homework") {
                    section.addPage("Biology");
                    section.addPage("Spanish");
                    section.addPage("Computer Science");
                }
            });
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

