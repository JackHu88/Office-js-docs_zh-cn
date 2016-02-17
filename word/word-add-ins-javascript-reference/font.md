# Font 对象（适用于 Word 的 JavaScript API）

表示字体。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
| 属性   | 类型|说明
|:---------------|:--------|:----------|
|bold|bool|获取或设置表示字体是否为粗体的值。如果字体格式为粗体则为 true，否则为 false。|
|color|string|获取或设置指定字体的颜色。您可以提供“#RRGGBB”格式的值或其他颜色名称。|
|doubleStrikeThrough|bool|获取或设置表示字体是否具有双删除线的值。如果字体格式设置为加双删除线的文本则为 true，否则为 false。|
|highlightColor|string|获取或设置指定字体的突出显示的颜色。您可以提供“#RRGGBB”格式的值或其他颜色名称。|
|italic|bool|获取或设置表示字体是否为斜体的值。如果字体为斜体则为 true，否则为 false。|
|name|string|获取或设置表示字体名称的值。|
|strikeThrough|bool|获取或设置表示字体是否具有删除线的值。如果字体格式设置为加删除线的文本则为 true，否则为 false。|
|subscript|bool|获取或设置表示字体是否为下标的值。如果字体格式为下标则为 true，否则为 false。|
|superscript|bool|获取或设置表示字体是否为上标的值。如果字体格式为上标则为 true，否则为 false。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|size|**float**|获取或设置表示字体大小（以磅值表示）的值。|
|underline|[UnderlineType](underlinetype.md)|获取或设置表示字体的下划线类型的值。有效值为：“None”、“Single”、“Word”、“Double”、“Dotted”、“Hidden”、“Thick”、“Dashline”、“Dotline”、“DotDashLine”、“TwoDotDashLine” 和 “Wave”|

## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息

### load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
无效

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the font property for all of the paragraphs.
    context.load(paragraphs, 'font');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object for the font object on the first paragraph in the collection.
        var font = paragraphs.items[0].font;
        
        // Queue a set of property value changes on the font proxy object.
        font.size = 32;
        font.bold = true;
        font.color = '#0000ff';
        font.highlightColor = '#ffff00';
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('The font has changed.');
        });  
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## 属性访问示例

### 更改字体名称
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the current selection's font name.
    selection.font.name = 'Arial';
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font name has changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 更改字体颜色
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the font color of the current selection.
    selection.font.color = 'blue'; 
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font color of the selection has been changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 更改字体大小
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the current selection's font size.
    selection.font.size = 20;
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font size has changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 突出显示选定文本
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to highlight the current selection.
    selection.font.highlightColor = '#FFFF00'; // Yellow
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection has been highlighted.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 粗体格式文本
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to make the current selection bold.
    selection.font.bold = true;
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection is now bold.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### 加下划线格式文本
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.thick;
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### 加删除线格式文本
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to strikethrough the font of the current selection.
    selection.font.strikeThrough = true; 
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has a strikethrough.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## 支持详细信息

在运行时检查过程中使用[要求设置](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)。 
