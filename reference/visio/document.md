# <a name="document-object-javascript-api-for-visio"></a>Document 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_

表示 Document 类。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
| 关系 | 类型    |说明|
|:---------------|:--------|:----------|
|application|[Application](application.md)|表示包含此文档的 Visio 应用程序实例。只读。|
|pages|[PageCollection](pagecollection.md)|表示一组与文档相关联的页面。只读。|
|view|[DocumentView](documentview.md)|返回 DocumentView 对象。只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getActivePage()](#getactivepage)|[Page](page.md)|返回文档的活动页。|
|[load(param: object)](#loadparam-object)|无效|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[setActivePage(PageName: string)](#setactivepagepagename-string)|void|设置文档的活动页。|
|[startDataRefresh()](#startdatarefresh)|void|触发刷新关系图中所有页面的数据。|

## <a name="method-details"></a>方法详细信息


### <a name="getactivepage"></a>getActivePage()
返回文档的活动页。

#### <a name="syntax"></a>语法
```js
documentObject.getActivePage();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var activePage = document.getActivePage();
    activePage.load();
    return ctx.sync().then(function () {
    console.log("pageName: " +activePage.name);
      });   
}).catch(function(error) {
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
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void

### <a name="setactivepagepagename-string"></a>setActivePage(PageName: string)
设置文档的活动页。

#### <a name="syntax"></a>语法
```js
documentObject.setActivePage(PageName);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|PageName|string|页面的名称|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var pageName = "Page-1";
    document.setActivePage(pageName);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="startdatarefresh"></a>startDataRefresh()
触发刷新关系图中所有页面的数据。

#### <a name="syntax"></a>语法
```js
documentObject.startDataRefresh();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    document.startDataRefresh();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```### Property access examples
```js
Visio.run(function (ctx) { 
    var pages = ctx.document.pages;
    var pageCount = pages.getCount();
    return ctx.sync().then(function () {
        console.log("Pages Count: " +pageCount.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var documentView = ctx.document.view;
    documentView.disableHyperlinks();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var application = ctx.document.application;
    application.showToolbars = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

