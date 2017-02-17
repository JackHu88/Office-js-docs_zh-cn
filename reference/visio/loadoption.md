# <a name="object-load-options-javascript-api-for-visio"></a>对象加载选项（适用于 Visio 的 JavaScript API）

>**注意：**Visio JavaScript API 暂处于预览阶段，可能会发生变更。暂不支持在生产环境中使用 Visio JavaScript API。

表示可以传递到加载方法，以指定在执行 **sync()** 方法时要加载的一组属性和关系的对象。sync() 方法可在 Visio 对象与相应的 JavaScript 代理对象之间同步状态。这会获取诸如选择、展开参数之类的选项，以指定要在对象上加载的一组属性，同时还允许对集合进行分页。

它还可用于提供包含要加载的属性和关系的字符串，或提供包含要加载的属性和关系列表的数组。请参阅以下示例。

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## <a name="properties"></a>属性

| 属性 | 类型  | 说明 |
|:---------|:------|:------------|
|select    |object |提供在执行 executeAsync 调用时要加载的参数/关系名称的逗号分隔列表或数组，例如 "property1, relationship1"、[ "property1", "relationship1"]。可选。|
|expand    |object |提供在执行 executeAsync 调用时要加载的关系名称的逗号分隔列表或数组，例如 "relationship1, relationship2"、[ "relationship1", "relationship2"]。可选。|
|top       |int    |指定要包括在结果中的查询集合中的项数目。可选。|
|skip      |int    |指定要跳过且不包含在结果中的集合中的项数目。如果指定 **top**，跳过指定数目的项目后将开始选择结果。可选。|

