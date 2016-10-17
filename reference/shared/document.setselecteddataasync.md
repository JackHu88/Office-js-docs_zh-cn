
# <a name="document.setselecteddataasync-method"></a>Document.setSelectedDataAsync 方法
将数据写入文档中的当前选择。

|||
|:-----|:-----|
|**主机：**Access、Excel、PowerPoint、Project、Word、Word Online|**外接程序类型：**内容、任务窗格|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**包含最后一次更改的版本**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## <a name="parameters"></a>参数

|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _data_|数据可以是以下任意数据类型：<ul><li><b>string</b> (Office.CoercionType.Text) - 仅适用于 Excel、Excel Online、PowerPoint、PowerPoint Online、Word 和 Word Online。</li><li>数组的 <b>array</b> (Office.CoercionType.Matrix) - 仅适用于 Excel、Word 和 Word Online。</li><li>[TableData](../../reference/shared/tabledata.md) (Office.CoercionType.Table) - 仅适用于 Access、Excel、Word 和 Word Online。</li><li><b>HTML</b> (Office.CoercionType.Html) - 仅适用于 Word 和 Word Online。</li><li><b>Office Open XML</b>  (Office.CoercionType.Ooxml) - 仅适用于 Word 和 Word Online。</li><li><b>Base64 编码图像流</b> (Office.CoercionType.Image) - 仅适用于 Excel、PowerPoint、Word 和 Word Online。</li></ul>|当前选择中要设置的数据。必需。|**包含更改的版本：**1.1。若要支持 Access 内容外接程序，必须有 **Selection** 要求集 1.1 或更高版本。若要支持设置图像数据，必须有 **ImageCoercion** 要求集 1.1 或更高版本。若要对其进行设置以用于应用程序激活，请使用：<br/><br/>`<Requirements>`<br/>&nbsp;&nbsp;`<Sets DefaultMinVersion="1.1">`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`<Set Name="ImageCoercion"/>`<br/>&nbsp;&nbsp;`</Sets>`<br/>`</Requirements>`<br/><br/>可通过以下代码完成运行时检测 ImageCoercion 功能：<br/><br/>`if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaImageCoercion();`<br/>`} else {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaOoxml();`<br/>`}`|
| _options_|**object**|指定一组 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。选项对象可能包含用于设置选项的下列属性：<br/><ul><li>coercionType (<b><a href="735eaab6-5e31-4bc2-add5-9d378900a31b.htm">CoercionType</a></b>) - 指定如何强制设置数据。如果未设置此选项，则使用默认的 coercionType 值 Office.CoercionType.Text。</li><li>tableOptions (<b>object</b>) - 对于插入的表格，为指定<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">表格格式选项</a>（例如标题行、总计行和带状行）的键值对列表。 </li><li>cellFormat (<b>object</b>) - 对于插入的表格，为指定列、行或单元格范围以及适用于该范围的<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">单元格格式</a>的键值对列表。 </li><li>imageLeft (<b>number</b>) - 此选项适用于插入图像。表示相对于 PowerPoint 幻灯片左侧的插入位置，以及相对于 Excel 中当前选定单元格的插入位置。Word 会忽略此值。此值以磅为单位。</li><li>imageTop (<b>number</b>) - 此选项适用于插入图像。表示相对于 PowerPoint 幻灯片顶部的插入位置，以及相对于 Excel 中当前选定单元格的插入位置。Word 会忽略此值。此值以磅为单位。</li><li>imageWidth (<b>number</b>) - 此选项适用于插入图像。表示图像宽度。如果提供此选项时未提供 imageHeight，那么图像会进行缩放，以匹配图像的宽度值。如果同时提供了图像的宽度和高度，那么图像会相应地调整大小。如果图像的高度或宽度均未提供，则会使用默认的图像大小和纵横比。此值以磅为单位。</li><li>imageHeight (<b>number</b>) - 此选项适用于插入图像。表示图像高度。如果提供此选项时未提供 imageWidth，那么图像会进行缩放，以匹配图像的高度值。如果同时提供了图像的宽度和高度，那么图像会相应地调整大小。如果图像的高度或宽度均未提供，则会使用默认的图像大小和纵横比。此值以磅为单位。</li><li>asyncContext (<b>object \| value</b>) - 用户定义的对象，适用于 <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a> 对象的 asyncContext 属性。当回调是命名的函数时，使用此选项为 <b>AsyncResult</b> 提供对象或值。</li></ul>|_tableOptions_ 和 _cellFormat_ 选项已在 v1.1 中添加，并且在 Excel 2013 和 Excel Online 中受支持。<br/><br/>_imageLeft_ 和 _ImageTop_ 选项在 Excel 和 PowerPoint 中受支持。|
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递到 **setSelectedDataAsync** 方法的回调函数中，[AsyncResult.value](../../reference/shared/asyncresult.value.md) 属性始终返回 **undefined**，因为没有要检索的对象或数据。


## <a name="remarks"></a>注解

为 _data_ 传递的值包含要写入当前选定内容的数据。如果该值为：


-  **一个字符串：**将插入可以强制为 **string** 的纯文本或任何文本。
    
    
    
    在 Excel 中，还可以将 _data_ 指定为有效公式，将该公式添加到选定的单元格。例如，将 _data_ 设置为 `"=SUM(A1:A5)"` 将计算指定范围中值的总数。但是，当在绑定单元格中设置公式时，设置后将无法从绑定单元格读取添加的公式（或任何已有公式）。如果在选定的单元格上调用 [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法以读取其数据，方法可能仅返回在单元格中显示的数据（即公式的结果）。
    
-  **数组的数组（“矩阵”）：**将插入不带标题的表数据。例如，若要将数据写入三行两列，可以传递如下的数组：`[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`。若要将数据写入三行一列，可以传递如下的数组：`[["R1C1"], ["R2C1"], ["R3C1"]]`。
    
    
    
    在 Excel 中，还可以将 _data_ 指定为数组的数组，其中包含有效公式以将其添加到选定单元格。例如，如果不覆盖任何其他数据，将 _data_ 设置为 `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` 会将这两个公式添加到选区。与在单个单元格上将公式设置为“text”一样，设置后你将无法读取添加的公式（或任何已有公式），只能读取公式的结果。
    
-  **[TableData](../../reference/shared/tabledata.md) 对象：**将插入带标题的表格。
    
    
    
     **注意：**在 Excel 中，如果您在为 _data_ 参数传递的 **TableData** 对象中指定公式，您可能无法获取预期的结果，这是由于 Excel 的“计算列”功能，此功能会自动复制列中的公式。要解决此问题，如果您希望将包含公式的 _data_ 写入到选定表中，请尝试将数据指定为数组的数组（而非 **TableData** 对象），并将 _coercionType_ 指定为 **Microsoft.Office.Matrix** 或“矩阵”。
    
 **特定于应用程序的行为**

此外，将数据写入选择时以下特定于应用程序的操作适用。

 **对于 Word**


- 如果未做出任何选择且插入点位于有效位置，则将在如下的插入点插入指定的 _data_：
    
      - 如果 _data_ 为字符串，则插入指定的文本。
    
  - 如果 _data_ 是数组的数组（“矩阵”）或 **TableData** 对象，则插入新的 Word 表。
    
  - 如果 _data_ 为 HTML，则插入指定的 HTML。
    
     >**重要说明**：如果你插入的 HTML 都是无效的，则 Word 不会引发错误。Word 会尽可能多地插入 HTML，并且省略任何无效数据。
  - 如果 _data_ 为 Office Open XML，则插入指定的 XML。
    
  - 如果 _data_ 为 base64 编码的图像流，则插入指定的图像。
    
- 如果做出了选择，则会按照上述相同的规则替换为指定的 _data_。
    
-  **插入图像**：插入的图像以嵌入式形式存放。**imageLeft** 和 **imageTop** 参数将被忽略。图像的纵横比始终被锁定。如果只给定了 **imageWidth** 参数或只给定了 **imageHeight**参数，另一个值将自动扩展以保留原始纵横比。
    
 **对于 Excel**


- 如果选择了一个单元格：
    
      - 如果 _data_ 为字符串，则将指定文本作为当前单元格的值插入。
    
  - 如果 _data_ 是一系列数组（“矩阵”），则插入一组指定的行和列（如果不会覆盖周围单元格中的其他任何数据的话）。
    
  - 如果 _data_ 为 **TableData** 对象，则插入具有一组指定的行和标题的新 Excel 表（如果不会覆盖周围单元格中的其他任何数据的话）。
    
- 如果选择了多个单元格且形状与 _data_ 的形状不一致，则会返回错误。
    
- 如果选择了多个单元格且选定内容的形状与 _data_ 的形状完全一致，则选定单元格的值会根据 _data_ 中的值进行更新。
    
-  **插入图像**：插入的图像为浮动图像。**imageLeft** 和 **imageTop** 位置参数是相对于当前选定的单元格而言的。允许 **imageLeft** 和 **imageTop** 的值为负数，将由 Excel 重新调整在工作表内放置的图像位置。除非 **imageWidth** 和 **imageHeight** 参数均已提供，否则将锁定图像的纵横比。如果只给定 **imageWidth** 参数或只给定 **imageHeight** 参数，则另一个值将自动扩展以保留原始纵横比。
    
在其他所有情况下，都会返回错误。

 **对于 Excel Online**

除了上面介绍的针对 Excel 的行为外，在 Excel Online 中写入数据时还需遵循下列限制。 


- 调用一次此方法时，使用 _data_ 参数写入工作表的单元格总数不得超过 20,000 个。
    
- 传递给 _cellFormat_ 参数的_格式设置组_的数量不能超过 100。每个格式设置组由应用于特定单元格范围的一组格式组成。例如，以下调用向 _cellFormat_ 传递了两个格式设置组。
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

 **对于 PowerPoint**

插入的图像为浮动图像。**imageLeft** 和 **imageTop** 位置参数是可选的，但如果已提供，应同时提供。如果提供单个值，则它将被忽略。允许 **imageLeft** 和 **imageTop** 的值为负数，且可以将图像放置在幻灯片的外部。如果未给定任何可选参数，并且幻灯片有一个占位符，则图像将替换幻灯片中的占位符。除非 **imageWidth** 和 **imageHeight** 参数均已提供，否则将锁定图像纵横比。如果只给定 **imageWidth** 参数或只给定 **imageHeight** 参数，则另一个值将自动扩展以保留原始纵横比。


## <a name="example"></a>示例

以下示例将选定文本或单元格设置为“Hello World!”。如果失败，则显示 [error.message](../../reference/shared/error.message.md) 属性的值。


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



指定可选 _coercionType_ 参数可让你指定你想要写入选择中的数据类型。以下示例将数据作为一个两列三行的数组写入，将 _coercionType_ 指定为该数据结构的 `"matrix"`，并且如果失败，则显示 [error.message](../../reference/shared/error.message.md) 属性的值。




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



以下示例将数据作为带标题且包含一列四行的表写入，将 _coercionType_ 指定为相应数据结构的 `"table"`。如果失败，则显示 [error.message](../../reference/shared/error.message.md) 属性的值。




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 在 Word 中，如果你要将 HTML 写入选区，你可以将 _coercionType_ 参数指定为以下示例中所示的 `"html"`，它使用 HTML `<b>` 标记以粗体显示“Hello”。




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在 Word、PowerPoint 或 Excel 中，如果你要将图像写入选区，可以将 _coercionType_ 参数指定为 `"image"`，如以下示例中所示。请注意，Word 会忽略 imageLeft 和 imageTop。




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## <a name="support-details"></a>支持详细信息


以下矩阵中的选中标记（![对勾符号](../../images/mod_off15_checkmark.png)）指示此方法在相应的 Office 主机应用程序中受到支持。空单元格指示 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**

||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![对勾符号](../../images/mod_off15_checkmark.png)|||
|**Excel**|![对勾符号](../../images/mod_off15_checkmark.png)|![对勾符号](../../images/mod_off15_checkmark.png)|![对勾符号](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![对勾符号](../../images/mod_off15_checkmark.png)|![对勾符号](../../images/mod_off15_checkmark.png)|![对勾符号](../../images/mod_off15_checkmark.png)|
|**Word**|![对勾符号](../../images/mod_off15_checkmark.png)|![对勾符号](../../images/mod_off15_checkmark.png)|![对勾符号](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**在要求集中可用**|Selection|
|**最低权限级别**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|在 Word 和 Word Online 中，现已开始支持将数据编写为 base64 编码的图像流。|
|1.1|在 Word Online 中，现已开始支持将 _data_ 编写为**数组**的数组（矩阵）和 **TableData**（表）。|
|1.1|在 Office for iPad 的 Excel、PowerPoint 和 Word 中，支持与 Windows 桌面上的 Excel、PowerPoint 和 Word 同等级别。|
|1.1|在 Word Online 中，现已开始支持将 _data_ 编写为 **string**（文本）。|
|1.1|添加使用可选参数 _tableOptions_ 和 _cellFormat_ 在 Excel 相关外接程序中 [插入表时对设置格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)的支持。|
|1.1|增加了对在 Access 相关外接程序中写入表数据的支持。|
|1.0|引入|
