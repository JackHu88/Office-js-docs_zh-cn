
# Document.getSelectedDataAsync 方法
读取包含在文档的当前选择中的数据。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint、Project、Word|
|**在要求集中可用**|Selection|
|**选择内容中的最后更改**|1.1|

```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>主机支持</b></td></tr><tr><td><b>Office.CoercionType.Text</b>（字符串）</td><td>仅限 Excel、Excel Online、PowerPoint、PowerPoint Online、Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.Matrix</b>（数组的数组）</td><td>仅限 Excel、Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.Table</b>（[TableData](../../reference/shared/tabledata.md) 对象）</td><td>仅限 Access、Excel、Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>仅限 Word。</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>仅限 Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>仅限 PowerPoint 和 PowerPoint Online</td></tr></table>|要返回的数据结构的类型。必需。||
| _选项_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>指定返回结果的数字或日期值是经过格式化处理还是未经格式化处理。</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>指定在检索数据时是否应用筛选。可选。</td><td>在 Word 文档中忽略此参数。</td></tr><tr><td><i>asyncContext</i></td><td><b>array</b>、<b>boolean</b>、<b>null</b>、<b>number</b>、<b>object</b>、<b>string</b> 或 <b>undefined</b></td><td>在未经更改的 <b>AsyncResult</b> 对象中返回的任何类型的用户定义的项目。</td><td></td></tr></table>|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给  **getSelectedDataAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问当前选区中的值，这些值以您使用  _coercionType_ 参数指定的数据结构或者格式返回。（请参阅 **备注** ，了解有关数据强制的详细信息。）|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

在任务窗格应用程序或内容外接程序中，使用  **getSelectedDataAsync** 方法编写可从用户在文档、电子表格、演示文稿或项目中的选区中读取数据的脚本。例如，用户在一个 Word 文档中选择了内容后，您可以使用 **getSelectedDataAsync** 方法读取该选区，然后将其作为查询或其他操作提交给 Web 服务。

读取选区后，您还可以使用 [Document](../../reference/shared/document.setselecteddataasync.md) 对象的 [setSelectedDataAsync](../../reference/shared/document.addhandlerasync.md) 和 **addHandlerAsync** 方法[写回选区或添加事件处理程序](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)，以检测用户是否更改了该选区。

只要选区是活动的，**getSelectedDataAsync** 方法便可读取其内容。在 Word 和 Excel 外接程序中，如果您需要建立对用户选区的持久读写关联，则请改用 [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) 方法[绑定到该选区](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。

使用 **getSelectedDataAsync** 方法的 _coercionType_ 参数来指定读取的选定数据的数据结构或格式。



|**指定的 _coercionType_**|**返回的数据**|**Office 主机应用程序支持**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** 或 `"text"`|一个字符串。|Word、Excel、PowerPoint 和 Project。<br/><br/> **注意**：在 Excel 中，即使某单元格的一个子集被选中，仍然返回全部单元格的内容。|
|**Office.CoercionType.Matrix** 或 `"matrix"`|数组的数组。例如，` [['a','b'], ['c','d']]` 为两行两列的选区。|Word 和 Excel。|
|**Office.CoercionType.Table** 或 `"table"`|[TableData](../../reference/shared/tabledata.md) 对象用于读取带标题的表格。|Word 和 Excel。|
|**Office.CoercionType.Html** 或 `"html"`|以 HTML 格式。|仅限 Word。|
|**Office.CoercionType.Ooxml** 或 `"ooxml"`|以 Open Office XML (OpenXML) 格式。|仅限 Word。<br/><br/> **提示**：当开发外接程序代码时，可以使用 _getSelectedDataAsync_ 方法的 `"ooxml"`**coercionType** 查看您在 Word 文档中选择的内容如何被定义为 OpenXML 标签。然后，使用 [Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) 方法的数据参数中的标签将带有格式或结构的内容写入文档。例如，您可以[将图像插入文档](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx)作为 OpenXML。|
|**Office.CoercionType.SlideRange** 或“slideRange”|一个包含名为“幻灯片”的数组 JSON 对象，该数组包含所选幻灯片的 ID、标题和索引。  **注意：**要选择多张幻灯片，用户必须在“**普通**”、“**大纲视图**”或“**幻灯片排序**”视图中编辑演示文稿。 此外，此方法在“**母版视图**”中不受支持。例如，`{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` 对应于两张幻灯片的选区。|仅限 PowerPoint。|
如果选区的数据结构与指定  _coercionType_ 不匹配，则 **getSelectedDataAsync** 方法将尝试强制将数据转换到该类型或结构。如果选区无法强制转换为指定的 **Office.CoercionType**，则  **AsyncResult.status** 属性将返回 `"failed"`。


## 示例

若要读取当前选区的值，您需要编写一个可读取该选区的回调函数。如下示例显示如何操作：


-  **将可读取当前选区值的回调函数传递给** _getSelectedDataAsync_ 方法的 **callback** 参数。
    
-  **将该选区读取为**未格式化且未经过筛选的文本。
    
-  在外接程序页面上 **显示该值**。
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
            }            
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|Selection|
|**最低权限级别**|[ReadDocument（需要使用 ReadAllDocument 来获得 Office Open XML）](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1| 在 Word Online 中，增加了对将 **Office.CoercionType.Matrix** 和 **Office.CoercionType.Table** 作为 _coercionType_ 参数的支持。|
|1.1|在 Office for iPad 的 Excel、PowerPoint 和 Word 中，增加了与在 Windows 桌面上 Excel、PowerPoint 和 Word 相同的支持级别。|
|1.1| 在 Word Online 中，增加了对将 **Office.CoercionType.Text** 作为 _coercionType_ 参数的支持。|
|1.1|在 PowerPoint 相关内容外接程序中，您可以获取所选幻灯片范围的 ID、标题和索引，通过将 **Office.CoercionType.SlideRange** 传递为 _getSelectedDataAsync_ 方法的 **coercionType** 参数实现。请参阅 [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) 方法主题，获取使用此值导航当前选中幻灯片的方法示例。|
|1.0|引入|
