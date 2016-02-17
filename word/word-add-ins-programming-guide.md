# Word 外接程序编程概述

_适用于：Word 2016、Word for iPad、Word for Mac_

Word 2016 引入了一种用于 Word 对象的新对象模型。此对象模型是 Office.js 除现有对象模型之外提供的，用于创建 Word 外接程序。此对象模型可通过由 Web 应用程序托管的 JavaScript 进行访问。

## 指令清单

新的 Word 外接程序 JavaScript API 使用与 Office 2013 外接程序模型相同的清单格式。清单描述外接程序在何处托管、如何显示、权限及其他信息。了解有关如何自定义[外接程序清单](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx)的详细信息。 

您具有多个发布 Word 外接程序清单的选项。了解如何[将 Office 外接程序发布](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx)到网络共享、应用程序目录或 Office 商店。

## 了解适用于 Word 的 JavaScript API

适用于 Word 的 JavaScript API 通过 Office.js 加载。它提供一组 JavaScript 代理对象，这些对象用于对使用 Word 文档内容的一组命令进行排队。这些命令可以批量运行。批量运行的结果是对 Word 文档采取的操作，例如插入内容、将 Word 对象与 JavaScript 代理对象同步。 

### 运行外接程序

让我们来看看运行外接程序时需要什么。所有外接程序都应该有一个 Office.initialize 事件处理程序。阅读[了解 API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx)，获取有关外接程序初始化的详细信息。  

Word 外接程序通过向 Word.run() 方法传递函数来执行。传递到运行方法的函数必须具有上下文参数。此[上下文对象](word-add-ins-javascript-reference/requestcontext.md)不同于您从 Office 对象获取的上下文对象，尽管它用于与 Word 运行时环境交互的相同目的。此上下文对象提供了对 Word JavaScript 对象模型的访问。让我们来看看基本 Word 外接程序的评论和代码：

**示例 1.Word 外接程序的初始化和执行**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

示例 1 说明了创建 Word 外接程序所需的基本代码。它会初始化 Office.js，并且包含用于与 Word 文档交互的运行方法。

### 代理对象

Word JavaScript 对象模型与 Word 中的对象为松散耦合。Word JavaScript 对象是用于 Word 文档中的真实对象的代理对象。对代理对象执行的所有操作都不会在 Word 中实现，Word 文档的状态不会在代理对象中实现，直至文档状态已同步。运行 context.sync() 时将同步文档状态。sync() 方法主要运行每个代理对象的队列中的命令集。示例 2 说明如何创建代理 body 对象以及用于在代理 body 对象上加载文本属性的排队命令，然后将 Word 文档中的正文与 body 代理对象同步。 

**示例 2.将文档正文与 body 代理对象同步。**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values. 
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

### 命令队列

Word 代理对象具有访问和更新对象模型的方法。这些方法按其在批处理中排队的顺序依次执行。在调用 context.sync() 之前会形成一批命令。将会执行在使用上下文的所有对象中排队的所有命令。  

在示例 3 中，我们演示了命令队列的工作原理。调用 context.sync() 时，发生的第一件事是[加载正文文本的命令](Word%20Add-ins%20JavaScript%20Reference/loadoption.md)在 Word 中执行。然后，将执行在 Word 上的正文中插入文本的命令。结果将返回到 body 代理对象。Word JavaScript 中的 body.text 属性的值为文本插入到 Word 文档<u>之前</u> Word 文档正文的值。 

**示例 3.执行命令批处理。**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

## 向我们提供反馈

您的反馈对我们意义重大。 

* 查看文档并在此存储库中直接[提交问题](https://github.com/OfficeDev/office-js-docs/issues)，告诉我们您在其中发现的任何疑问和问题。
* 让我们了解您的编程体验、您希望在未来版本中看到的功能、代码示例，等等。请在[此网站](http://officespdev.uservoice.com/)输入您的建议和想法。


## 其他资源

* [Word 外接程序](word-add-ins.md)
* [Word 外接程序 JavaScript 参考](word-add-ins-javascript-reference.md)
* [Office 外接程序](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [开始使用 Office 外接程序](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;GitHub 上的 Word 外接程序&lt;/a&gt;
* [适用于 Word 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)

