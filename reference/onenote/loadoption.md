# 对象加载选项 

表示一个对象，该对象可以传递到加载方法，以指定在执行 sync() 方法时将加载的属性和关系集。sync() 方法可将外接程序中 OneNote 对象与相应的 JavaScript 代理对象之间的状态同步。这会将选择、展开参数等选项考虑在内，以指定要在对象上加载的属性集，并允许对集合进行分页。

它还可用于提供包含要加载的属性和关系的字符串，或提供包含要加载的属性和关系列表的数组。请参阅以下示例。

```js   
object.load('<var1>,<relationship1/var2>');

// Pass the parameter as an array.
object.load(["var1", "relationship1/var2"]);
```

## 属性
| 属性     | 类型   |说明|
|:---------------|:--------|:----------|
|select|object|提供在执行 sync 调用时要加载的参数/关系名称的逗号分隔列表或数组，例如 "property1, relationship1"、[ "property1", "relationship1"]。可选。|
|expand|object|提供在执行 sync 调用时要加载的关系名称的逗号分隔列表或数组，例如 "relationship1, relationship2"、[ "relationship1", "relationship2"]。可选。|
|top|int|指定要包括在结果中的查询集合中的项数目。可选。|
|skip|int|指定要跳过且不包含在结果中的集合中的项数目。如果指定 `top`，跳过指定数目的项目后将开始选择结果。可选。|

#### 示例

在示例中，获取当前分区中的前五页的页面标题和缩进级别。

```js
OneNote.run(function (context) { 
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages.           
    pages.load({ "select":"title,pageLevel", "top":5, "skip":0 });
    return context.sync()
        .then(function() {
            
            // Iterate through the collection of pages.    
            $.each(pages.items, function(index, page) {
                
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Indentation level: " + page.pageLevel);
                
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
    });
```
