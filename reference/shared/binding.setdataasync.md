
# Binding.setDataAsync 方法
将数据写入指定的绑定对象表示的文档的绑定部分。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|MatrixBindings, TableBindings, TextBindings|
|**TableBindings 中的最后更改**|1.1|

```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>仅限 Excel、Excel Online、Word 和 Word Online</td></tr><tr><td><b>array</b>（数组的数组 –“矩阵”）</td><td>仅限 Excel 和 Word</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp161002">
  <b>TableData</b></a></td><td>仅限 Access、Excel 和 Word</td></tr><tr><td><b>HTML</b></td><td>仅限 Word 和 Word Online</td></tr><tr><td><b>Office Open XML</b></td><td>仅限 Word</td></tr></table>|当前选择中要设置的数据。可选。必需。|**在其中所做的更改：**1.1。对 Access 相关内容外接程序的支持要求 **TableBinding** 要求集 1.1 或更高版本。|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|指定如何强制设置数据。 ||
| _columns_|**字符串数组**| 指定列名称。|**在其中添加：**v1.1。仅针对 Access 相关内容外接程序中的表绑定。|
| _rows_|**Office.TableRange.ThisRow**|指定预定义的字符串“thisRow”以设置当前所选行中的数据。 |**在其中添加：**v1.1。仅针对 Access 相关内容外接程序中的表绑定。|
| _startColumn_|**number**|为数据子集指定基于零的起始列。 |仅针对表或矩阵绑定。如果省略，数据设置为从第一列开始。|
| _startRow_|**number**|为绑定中的数据子集指定基于零的开始行。 |仅针对表或矩阵绑定。如果省略，数据设置为从第一行开始。|
| _tableOptions_|**object**|对于插入的表格，为指定[表格式选项](../../docs/excel/format-tables-in-add-ins-for-excel.md)（例如标题行、总计行和带状行）的键值对列表。 |**在以下版本中添加：**v1.1。**受以下版本支持：**Excel。|
| _cellFormat_|**object**|对于插入的表，为指定列、行或单元格范围以及适用于该范围的[单元格格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)的键值对列表。|**在 v1.1 中添加**。**受以下版本支持：**Excel 和 Excel Online。|
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **setDataAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined**，这是因为没有要检索的对象或数据。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

为 _data_ 传递的值包含要写入绑定中的数据。传递的值的类型决定将要写入的内容，如下表中所示。



|**_data_ 值**|**写入的数据**|
|:-----|:-----|
|一个**字符串**|将写入可以强制为 **string** 的纯文本或任何文本。|
|数组的数组（“矩阵”）|将写入不带标题的表格数据。例如，若要将数据写入三行两列中，可以像这样传递任何数组：` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`。若要写入一列的三行，像这样传递数组：`[["R1C1"], ["R2C1"], ["R3C1"]]`|
|[TableData](../../reference/shared/tabledata.md) 对象|将写入带标题的表格数据。|
此外，这些特定于应用程序的操作在将数据写入绑定时适用。

 **对于 Word**，按如下所示将指定的 _data_ 写入绑定：



|**_data_ 值**|**写入的数据**|
|:-----|:-----|
|一个**字符串**|写入指定的文本。|
|数组的数组（“矩阵”）或 **TableData** 对象|写入 Word 表格。|
|HTML|写入指定的 HTML。
 >**重要信息**  如果您写入的任何 HTML 均无效，则 Word 将不会引发错误。Word 将尽可能多的写入 HTML，因为这可以并且将省略任何无效数据。

|
|Office Open XML ("Open XML")|写入指定的 XML。|  **对于 Excel**，按如下所示，指定的 _data_ 写入绑定：



|**_data_ 值**|**写入的数据**|
|:-----|:-----|
|一个**字符串**|将文本插入为第一个绑定单元格的值。您还可以指定一个有效公式，将该公式添加到绑定单元格。例如，将 _data_ 设置为 `"=SUM(A1:A5)"` 将计算指定范围中值的总数。但是，当您在绑定单元格中设置公式时，设置后将无法从绑定单元格读取添加的公式（或任何已有公式）。如果您在绑定单元格上调用 [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) 方法以读取其数据，该方法可能仅返回在单元格中显示的数据（即公式的结果）。|
|数组的数组（“矩阵”），形状与指定绑定的形状完全匹配|写入一组行或列。您还可以指定数组的数组，其中包含用于将其添加到绑定单元格的有效公式。例如，将 _data_ 设置为 `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` 会将这两个公式添加到包含两个单元格的绑定。与在单个绑定单元格上设置公式一样，您无法使用 **Binding.getDataAsync** 方法从绑定中读取添加的公式（或任何已有的公式）– 它仅返回在绑定单元格中显示的数据。|
|**TableData** 对象，并且表的形状与绑定表匹配。|如果不会覆盖周围单元格中的其他数据，则写入一组行和/或标题。**注意：**如果您在为 **data** 参数传递的 _TableData_ 对象中指定公式，您可能无法获取预期的结果，这是由于 Excel 的“计算列”功能，此功能会自动复制列中的公式。要解决此问题，如果您希望将包含公式的 _data_ 写入到绑定表中，请尝试将数据指定为数组的数组（而非 **TableData** 对象），并将 _coercionType_ 指定为 **Microsoft.Office.Matrix** 或“矩阵”。|
 **Excel Online 的其他标记**


- 对此方法的单个调用中，传递给 _data_ 参数的值中的单元格总数不能超过 20,000。
    
- 传递给 _cellFormat_ 参数的_格式设置组_的数量不能超过 100。每个格式设置组由应用于特定单元格范围的一组格式组成。例如，以下调用向 _cellFormat_ 传递了两个格式设置组。
    
```js
  Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

在其他所有情况下，都会返回错误。

如果指定了可选 **startRow** 和 _startColumn_ 参数，并且它们指定一个有效的范围，则 _setDataAsync_ 方法将在表子集或矩阵绑定中写入数据。


## 示例




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

指定可选 _coercionType_ 参数可让你指定你想要写入绑定中的数据类型。 例如，在 Word 中，如果你想要将 HTML 写入文本绑定，你可以将 _coercionType_ 参数指定为 `"html"`，如以下示例所示，它将使用 HTML `<b>` 标记以粗体形式显示“Hello”。




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在本示例中，对 **setDataAsync** 的调用将 _data_ 参数作为数组的数组传递（以创建一列的三行），并使用 _coercionType_ 参数将数据结构指定为 `"matrix"`。




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在本示例的 `writeBoundDataTable` 函数中，对 **setDataAsync** 的调用会将 _data_ 参数作为 **TableData** 对象传递（写入三列和三行），并使用 _coercionType_ 参数将数据结构指定为 `"table"`。 

在 `updateTableData` 函数中，对 **setDataAsync** 的再次调用会将 _data_ 参数作为 **TableData** 对象传递（但作为使用新标题和三行的单个列），以更新表格最后一列由 `writeBoundDataTable` 函数创建的值。从零开始的可选 _startColumn_ 参数指定为 2，以替换表格第三列中的值。




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
         }     
     });   
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
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|MatrixBindings, TableBindings, TextBindings|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|<ul><li>在 Access 相关外接程序中，添加了对写入表格数据的支持。</li><li>在 Excel 相关外接程序中，添加了使用 <span class="parameter" sdata="paramReference">tableOptions</span> 和 <span class="parameter" sdata="paramReference">cellFormat</span> 可选参数<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">将数据写入表格绑定时对设置格式</a>的支持。</li></ul>|
|1.0|引入|

## 另请参阅



#### 其他资源


[绑定到文档或电子表格中的区域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
