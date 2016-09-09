# PageContent 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示在页面上包含顶级内容类型的地区，例如 Outline 或 Image。可对 PageContent 对象分配一个 XY 位置。

## 属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|string|获取 PageContent 对象的 ID。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|左边|double|获取或设置 PageContent 对象的左边（X 轴）位置。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|顶部|double|获取或设置 PageContent 对象的顶部（Y 轴）位置。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|类型|string|获取 PageContent 对象的类型。 只读。 可能的值是：边框、图像、其他。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## 关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|图像|[Image](image.md)|获取 PageContent 对象中的 Image。 如果 PageContentType 不为 Image，则引发异常。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|墨迹|[FloatingInk](floatingink.md)|获取 PageContent 对象中的 Ink。 如果 PageContentType 不为 Ink，则引发异常。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|边框|[分级显示](outline.md)|获取 PageContent 对象中的 Outline。 如果 PageContentType 不为 Outline，则引发异常。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[页面](page.md)|获取包含 PageContent 对象的页面。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## 方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|删除 PageContent 对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## 方法详细信息


### delete()
删除 PageContent 对象。

#### 语法
```js
pageContentObject.delete();
```

#### 参数
无

#### 返回
void

#### 示例
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;

    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
                return context.sync();
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
