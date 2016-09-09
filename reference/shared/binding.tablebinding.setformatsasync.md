
# TableBinding.setFormatsAsync 方法
设置或更新绑定表中指定项目和数据的格式。

|||
|:-----|:-----|
|**主机：**|Excel|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|不在集合中|
|**在其中添加**|1.1|

```
bindingObj.setFormatsAsync(cellFormat [,options] , callback);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _cellFormat_|**array**|包含指定针对哪些单元格以及适用格式的一个或多个 JavaScript 对象的数组。必需。||
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给  **goToByIdAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined** 因为在设置格式时没有要检索的数据或对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 注解

 **指定 cellFormat 参数**

使用 _cellFormat_ 参数设置或更改单元格格式化值，例如宽度、高度、字体、背景、对齐等等。您作为 _cellFormat_ 参数传递的值是包含一个或多个 JavaScript 对象的列表的 **array**，可指定针对的单元格 (`cells:`) 以及适用的格式 (`format:`)。

_cellFormat_ 数组中每个 JavaScript 对象的格式为：

 `{cells:{`_cell_range_`}, format:{`_format_definition_`}}`

`cells:` 属性指定您希望使用以下值格式化的范围：


**单元格属性中支持的范围**


|**单元格范围设置**|**说明**|
|:-----|:-----|
| `{row: i}`|指定延伸到表中第 i 行数据的范围。|
| `{column: i}`|指定延伸到表中第 i 列数据的范围。|
| `{row: i, column: j}`|指定表中第 i 行到第 j 列数据的单元格范围。|
| `Office.Table.All`|指定整个表格，包括列标题、数据和总数（如果有）。|
| `Office.Table.Data`|仅指定表中的数据（不含标题和总数）。|
| `Office.Table.Headers`|仅指定标题行。|


`format:` 属性指定对应于 Excel“**设置单元格格式**”对话框中可用的设置子集的值（右键单击选择“**设置单元格格式**”或依次选择“**开始**” > “**格式**” > “**设置单元格格式**”）。

您可以指定 `format:` 属性的值作为 JavaScript 对象文字中一个或多个_属性名_ - _值_对列表。_属性名称_指定要设置的格式属性的名称，_值_则指定属性值。您可以为给定的格式指定多个值，如字体的颜色及大小。以下是三个 `format:` 属性值的示例：




```
//Set cells: font color to green and size to 15 points.
format: {fontColor : "green", fontSize : 15}
```




```
//Set cells: border to dotted blue.
format: {borderStyle: "dotted", borderColor: "blue"}
```




```
//Set cells: background to red and alignment to centered.
format: {backgroundColor: "red", alignHorizontal: "center"}
```

你可以通过指定 `numberFormat:` 属性中的数字格式 "code" 字符串，来指定数字格式。 你可以指定的数字格式字符串对应于你可以使用“**设置单元格格式**”对话框“**数字**”选项卡上的“**自定义**”类别在 Excel 中设置的字符串。 此示例演示如何将数字格式设置为带两个小数位的百分数形式：




```
format: {numberFormat:"0.00%"}
```

有关详细信息，请参阅如何 [创建自定义数字格式](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1#BM1)。



 **指定单个目标**

以下示例演示将标题行的字体颜色设置为红色的 _cellFormat_ 值。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: Office.Table.Headers, format: {fontColor: "red"}}], 
    function (asyncResult){});
```

 **指定多个目标**

**setFormatsAsync** 方法可支持在单个函数调用中对绑定表中的多个目标进行格式化。为此，您应为您要格式化的每个目标传递 _cellFormat_ 数组中的对象列表。例如，以下代码行将第一行的字体颜色设置为黄色，将第三行的第四个单元格设置为具有白色边框且使用粗体文本。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

若要在写入数据时设置表的格式，请使用 _Document.setSelectedDataAsync_ 或 _TableBinding.setDataAsync_ 方法的 [tableOptions](http://msdn.microsoft.com/library/4c1e13e9-b61a-47df-836c-3ca9aba4ca1c%28Office.15%29.aspx) 和 [cellFormat](http://msdn.microsoft.com/library/5b6ecf6f-c57f-4c0d-9605-59daee8fde13%28Office.15%29.aspx) 可选参数。

使用 **Document.setSelectedDataAsync** 和 **TableBinding.setDataAsync** 方法的可选参数设置格式仅适用于首次写入数据时设置格式的情况。若要在写入数据后更改格式，请使用下列方法：


- 若要更新单元格的格式（例如，字体颜色和样式），请使用 **TableBinding.setFormatsAsync** 方法（此方法）。
    
- 若要更新表选项（例如，带状行和筛选器按钮），请使用 [TableBinding.setTableOptions](../../reference/shared/binding.tablebinding.settableoptionsasync.md) 方法。
    
- 若要清除格式，请使用 [TableBinding.clearFormats](../../reference/shared/binding.tablebinding.clearformatsasync.md) 方法。
    
 **Excel Online 的其他标记**

传递给 _cellFormat_ 参数的_格式设置组_的数量不能超过 100。每个格式设置组由应用于特定单元格范围的一组格式组成。例如，以下调用向 _cellFormat_ 传递了两个格式设置组。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});

```

有关详细信息和示例，请参阅[如何设置 Excel 相关外接程序中表的格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**||**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|不在集合中。|
|**最低权限级别**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 的支持。|
|1.1|引入|
