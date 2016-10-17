# <a name="inkstrokepointer-object-(javascript-api-for-onenote)"></a>InkStrokePointer 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


对墨迹笔划对象及其内容父级的弱引用

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|contentId|string|表示此笔划所对应的页面内容对象的 ID|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-contentId)|
|inkStrokeId|string|表示墨迹笔划的 ID|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-inkStrokeId)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-load)|

## <a name="method-details"></a>方法详细信息


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
