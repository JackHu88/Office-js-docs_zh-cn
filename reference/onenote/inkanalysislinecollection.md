# <a name="inkanalysislinecollection-object-(javascript-api-for-onenote)"></a>InkAnalysisLineCollection 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示 InkAnalysisLine 对象的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|count|int|返回页面中的 InkAnalysisLines 数目。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|inkAnalysisLine 对象的集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisLine](inkanalysisline.md)|按其在集合中的 ID 或索引获取 InkAnalysisLine 对象。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|按其在集合中的位置获取 InkAnalysisLine。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
按其在集合中的 ID 或索引获取 InkAnalysisLine 对象。只读。

#### <a name="syntax"></a>语法
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number 或 string|InkAnalysisLine 对象的 ID 或集合中 InkAnalysisLine 对象的索引位置。|

#### <a name="returns"></a>返回
[InkAnalysisLine](inkanalysisline.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
按其在集合中的位置获取 InkAnalysisLine。

#### <a name="syntax"></a>语法
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[InkAnalysisLine](inkanalysisline.md)

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
