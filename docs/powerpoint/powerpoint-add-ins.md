
# <a name="create-content-and-task-pane-add-ins-for-powerpoint"></a>创建 PowerPoint 相关的内容和任务窗格外接程序

本文中的代码示例演示了开发 PowerPoint 内容加载项所使用的一些基本任务。这些示例都依赖于  `app.showNotification` 函数（此函数包含在 Visual StudioOffice 外接程序 项目模板中）来显示信息。如果您不使用 Visual Studio 开发加载项，则需要用自己的代码替换 `showNotification` 函数。其中某些示例还依赖于在这些函数作用域之外声明的 `globals` 对象： `var globals = {activeViewHandler:0, firstSlideId:0};`

这些代码示例要求您的项目 [引用 Office.js v1.1 库或更高版本](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。


## <a name="detect-the-presentation's-active-view-and-handle-the-activeviewchanged-event"></a>检测演示文稿的活动视图并处理 ActiveViewChanged 事件

`getFileView` 函数将调用 [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) 方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）视图。


```js
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

`registerActiveViewChanged` 函数将调用 [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) 方法，以注册 [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) 事件的处理程序。执行此函数后，当你更改演示文稿的视图时，`app.showNotification` 通知将显示活动视图模式（“阅读”或“编辑”）。




```js
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## <a name="get-the-url-of-the-presentation"></a>获取演示文稿的 URL

`getFileUrl` 函数将调用 [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) 方法以获取演示文稿文件的 URL。


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>导航到演示文稿中特定的幻灯片

`getSelectedRange` 函数将调用 [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法，以获取 `asyncResult.value` 返回的、包含名为“slides”的阵列的 JSON 对象，该阵列中包含所选幻灯片范围（或仅当前幻灯片）的 ID、标题和索引。它还会将所选范围内第一张幻灯片的 ID 保存到一个全局变量。


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

`goToFirstSlide` 函数将调用 [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) 方法，以转到上述 `getSelectedRange` 函数存储的第一张幻灯片的 ID。




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## <a name="navigate-between-slides-in-the-presentation"></a>在演示文稿的幻灯片之间导航

`goToSlideByIndex` 函数调用 **Document.goToByIdAsync** 方法以导航到演示文稿的下一个幻灯片。


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## <a name="additional-resources"></a>其他资源

- [如何按文档保留内容和任务窗格外接程序的外接程序状态和设置](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [读取数据并将其写入文档或电子表格中的活动选择区](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [从 PowerPoint 或 Word 相关外接程序中获取整个文档](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [在 PowerPoint 外接程序中使用文档主题](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
