# <a name="sectiongroupcollection-object-(javascript-api-for-onenote)"></a>SectionGroupCollection 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


代表分区组的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|count|int|返回集合中的分区组数。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-count)|
|items|[SectionGroup[]](sectiongroup.md)|sectionGroup 对象的集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-items)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|获取具有指定名称的分区组的集合。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[SectionGroup](sectiongroup.md)|按其在集合中的 ID 或索引获取分区组。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[SectionGroup](sectiongroup.md)|按其在集合中的位置获取分区组。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getbyname(name:-string)"></a>getByName(name: string)
获取具有指定名称的分区组的集合。

#### <a name="syntax"></a>语法
```js
sectionGroupCollectionObject.getByName(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|name|字符串|节组的名称。|

#### <a name="returns"></a>返回
[SectionGroupCollection](sectiongroupcollection.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("id"); 

    // Get the section groups with the specified name.
    var labsSectionGroups = sectionGroups.getByName("Labs");

    // Queue a command to load the section groups with the specified properties.
    labsSectionGroups.load("id,name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (labsSectionGroups.items.length > 0) {
                console.log("Section group name: " + labsSectionGroups.items[0].name);
                console.log("Section group ID: " + labsSectionGroups.items[0].id);
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
按其在集合中的 ID 或索引获取分区组。只读。

#### <a name="syntax"></a>语法
```js
sectionGroupCollectionObject.getItem(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number or string|分区组的 ID 或集合中的分区组的索引位置。|

#### <a name="returns"></a>返回
[SectionGroup](sectiongroup.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
按其在集合中的位置获取分区组。

#### <a name="syntax"></a>语法
```js
sectionGroupCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[SectionGroup](sectiongroup.md)

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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sectionGroups.items[0]
            $.each(sectionGroups.items, function(index, sectionGroup) {
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

