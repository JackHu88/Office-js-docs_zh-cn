
# Document.getFileAsync 方法
以高达 4194304 字节 (4 MB) 的切片形式返回整个文档文件。对于 iOS 外接程序而言，支持高达 65536 (64 KB) 的文件切片。请注意，若指定文件切片的大小上限超出允许限制，则会导致一个“内部错误”故障。 

|||
|:-----|:-----|
|**主机：**|Excel、PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|文件|
|**文件中的最后更改**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|指定返回文件的格式。必需。<br/><table><tr><th>主机</th><th>受支持的文件类型</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>Windows 桌面上的 PowerPoint</td><td>Office.FileType.Compressed、Office.FileType.Pdf</td></tr><tr><td>Windows 桌面、MAC 和 iPad 上的 Word</td><td>Office.FileType.Compressed、Office.FileType.Pdf 和 Office.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.Compressed、Office.FileType.Pdf 和 Office.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.Compressed、Office.FileType.Pdf</td></tr></table>|**在 1.1 中已更改**，请参阅[支持历史记录](#支持历史记录)|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _sliceSize_|**number**|指定最多 4194304 字节 (4 MB) 的所需切片大小（以字节为单位）。如果未指定，则使用 4194304 字节 (4 MB) 的默认切片大小。 ||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给  **getFileAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问 [File](../../reference/shared/file.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

对于运行在 Office 主机应用程序中（而非 Office for iOS 中）的外接程序， **getFileAsync** 方法支持以切片形式获取至多 4194304 字节 (4 MB) 的文件。对于运行在 iOS 的 Office 应用程序中的外接程序， **getFileAsync** 方法支持以切片形式获取至多 65536 (64 KB) 的文件。

可以使用以下枚举或文本值指定  _fileType_ 参数。


**FileType 枚举**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|以字节数组形式返回 Office Open XML (OOXML) 格式的完整文档（.docx、.pptx 或 .xslx）。|
|Office.FileType.Pdf|“pdf”|将 PDF 格式的整个文档作为字节数组返回。|
|Office.FileType.Text|"text"|只返回 **string**形式的文档文本。 |
内存中不允许两个以上的文档；否则 **getFileAsync** 操作将会失败。处理完文件后，使用 [File.closeAsync](../../reference/shared/file.closeasync.md) 方法关闭文件。


## 示例 - 获取 Office Open XML（"压缩"）格式的文档

以下示例以切片形式获取至多 65536 字节 (64 KB) 的 Office Open XML（"压缩"）格式的文档。注意：本示例中对  `app.showNotification` 的实现来自于 Office 外接程序的 Visual Studio 模板。


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## 示例 - 获取 PDF 格式的文档

下面的示例获取 PDF 格式的文档。


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|文件|
|**最低权限级别**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.1| 在 PowerPoint Online 中，增加了对将 **Office.FileType.Pdf** 作为 _fileType_ 参数的支持。|
|1.1| 在 PowerPoint Online 中，增加了对将 **Office.FileType.Compressed** 作为 _fileType_ 参数的支持。|
|1.1| 在 Word Online 中，增加了对将  **Office.FileType.Text** 作为 _fileType_ 参数的支持。|
|1.1| 在 Excel Online 中，增加了对将  **Office.FileType.Compressed** 作为 _fileType_ 参数的支持。|
|1.1| 在 Word Online 中，增加了对将  **Office.FileType.Compressed** 和 **Office.FileType.Pdf** 作为 _fileType_ 参数的支持。|
|1.1|在 Office for iPad 中的 PowerPointWord 上，增加了对将所有  **FileType** 值作为 _fileType_ 参数的支持。|
|1.1|在 Windows 桌面上的 Word 和 PowerPoint 中，增加了对将  **Office.FileType.Pdf** 作为 _fileType_ 参数的支持。|
|1.0|引入|
