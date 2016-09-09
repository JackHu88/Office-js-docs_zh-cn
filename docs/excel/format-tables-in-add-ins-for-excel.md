
# 设置 Excel 相关外接程序中表的格式


本文介绍了格式化 API 的不同功能并概述了如何使用这些功能。在本版本中，你可以编程方式指定仅适用于表格（而非 **Office.CoercionType.Text** 或 **Office.CoercionType.Matrix** 数据结构）且仅在 Excel 外接程序中的单元格格式化选项和其他选项。要设置外接程序格式化，请执行以下操作：

- 用户选择表（或者在其中以编程方式插入表的位置），然后你的加载项可以在该表上调用 **Document.setSelectedDataAsync** 方法以设置格式化。

- 如果工作簿已包含绑定表（或者你的加载项使用 [Bindings](../../reference/shared/bindings.bindings.md) 对象的“addFrom”方法之一在初始化时创建绑定表），你的加载项可以对这些绑定表调用 **Binding.setDataAsync** 方法以设置格式化。
    
>**重要说明：**若要使用这些新的和更新的方法在 Excel 外接程序中设置表格格式，外接程序项目必须 [使用或更新为使用 Office.js v1.1 或更高版本](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。

## 指定格式

要指定要设置的格式，可以使用包含一个或多个键值对的 JavaScript 对象文本。您可以在 JavaScript 对象的列表中组合一系列格式化设置。例如： 


```js
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

要应用格式化，请将 JavaScript 对象传递到支持表的格式化数据和其他功能的方法之一。

您可以通过两种方式使用格式化：


- 通过在传递到 [Document.setSelectedDataAysnc](../../reference/shared/document.setselecteddataasync.md) 或 [Binding.setDataAsync](../../reference/shared/binding.setdataasync.md) 方法的 _options_ 对象中指定可选的 _cellFormat_ 或 _tableOptions_ 参数，你的外接程序第一次将数据写入选择或绑定。
    
- 初始设置格式化之后，你可以使用专用于此目的的一个新方法 [清除或更新格式化](#清除或更新格式化)。
    

## 通过数据设置方法使用可选参数

对于表绑定，你可以在使用 **Document.setSelectedData** 或 **Binding.setDataAsync** 方法设置数据时，使用可选参数 _tableOptions_和 _cellFormat_ 指定格式。


### 可选参数 tableOptions

使用 _tableOptions_ 可选参数指定默认表格演示并启用或禁用部分表格功能，例如：**标题行**、**总计行**和**带状行** 你传递作为 _tableOptions_ 参数的值是一个包含键值对列表的 JavaScript 对象。 例如，


```js
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### 可选参数 cellFormat

使用 _cellFormat_ 可选参数更改单元格格式值，例如宽度、高度、字体、背景、对齐等。 你传递作为 _cellFormat_ 参数的值是一个数组，其中包含指定目标单元格和要对单元格应用的格式的 JavaScript 对象列表。 例如：


```js
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

你可以在 _cellFormat_ 数组中结合使用多个 `cells:` 和 `format:` 对，以最大限度地减少应用格式化所需的函数调用次数。


#### 单元格

使用 `cells:` 指定你希望对其应用格式化的列、行和单元格的范围。


**单元格值的支持范围**


|**单元格范围设置**|**说明**|
|:-----|:-----|
| `{row: i}`|指定延伸到表中第 i 行数据的范围。|
| `{column: i}`|指定延伸到表中第 i 列数据的范围。|
| `{row: i, column: j}`|指定表中第 i 行到第 j 列数据的单元格范围。|
| `Office.Table.All`|指定整个表格，包括列标题、数据和总数（如果有）。|
| `Office.Table.Data`|仅指定表中的数据（不含标题和总数）。|
| `Office.Table.Headers`|仅指定标题行。|

#### 格式

使用 `format:` 指定要应用到使用 `cells:` 定义为 JavaScript 键值对列表的范围的格式化。 有关支持值的列表，请参阅 [支持的格式化键和值](#支持的格式化键和值)。

 **指定 Excel Online 的格式限制**

设置 Excel Online 中的格式时，传递到 _cellFormat_ 参数的_格式组_ 的数量不能超过 100。 单个格式组包括一组应用于特定单元格范围的格式。 （换而言之，数组中的 `cells:` 对象文本之一中指定的全部内容都传递到了 _cellFormat_例如，以下调用向 _cellFormat_ 传递了两个格式组。




```js
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### 应用可选参数

在此版本中，仅 **Document.setSelectedDataAsync** 和 **TableBinding.setDataAsync** 方法支持在使用 _tableOptions_ 和 _cellFormat_ 可选参数的相同调用中写入数据并设置表格的格式。 在以下示例中，传递到各个方法的第一个参数（_data_ 参数）的 `tableData` 值必须为包含表格的定义和要写入的数据的 [TableData](../../reference/shared/tabledata.md)。

 **Document.setSelectedDataAsync 示例**




```js
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **TableBinding.setDataAsync 示例**




```js
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 >**注意：**调用 `Office.select("bindings#myBinding")` 假定工作表中已存在名为 `myBinding` 的绑定。


## 更新和清除格式化


通过 **Document.setSelectedDataAsync** 或 **TableBinding.setDataAsync** 方法的 _cellFormat_ 和 _tableOptions_ 可选参数设置格式时，仅会在首次调用时设置格式。 若要更新或清除格式，则必须使用 **TableBinding** 对象的三个新方法：**setFormatsAsync**、**setTableOptionsAsync** 和 **clearFormatsAsync**。


### 更新格式

[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md) 方法仅用于更新单元格格式，如宽度、高度、字体、背景和对齐。 它将“_cellFormat_”作为必填参数：


```js
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

[TableBinding.setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md) 方法仅用于更新表选项，如带状行和筛选器按钮。 它将 _tableOptions_ 作为必填参数：




```js
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### 清除格式

[TableBinding.clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md) 方法用于清除表格中的所有格式。 它需要 _asyncContext_ 可选参数和可选回调函数：


```js
Office.select("bindings#myBinding").clearFormatsAsync();
```


## 支持的格式键和值


下表列出了可以传递到 _cellFormat_ 或 _tableOptions_ 参数的支持键值对。

对于 `format:` 值，可用设置对应于“**格式单元格**”对话框（右击>“**格式单元格**”或功能区的“**主页**”选项卡上的“**格式**” > “**格式单元格**”）其中的一个子集。 对于 `tableOptions:` 值，设置对应于“**表格样式选项**”和功能区的“**表格工具**” |“**设计**”选项卡上的“**表格样式**”组。


 >**重要说明**：格式化 API 的方法仅支持下面汇总的选项和值。 如果指定这些以外的格式化选项或值，则处理行为未定义。 这些未定义的处理行为在支持的平台之间不一定一致；你不应基于任何特定平台的这些未定义行为的任何副作用开发外接程序。 但是，未定义的处理行为不应危害外接程序或与其交互的文档的状态和 UI。


**对齐**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `alignHorizontal:`|“general”\| "left" \| "center" \| "right" \| "fill" \| "justify" \| "center across selection" \| "distributed"|与缩进值结合使用时，仅支持以下组合：<br/><br/><ul><li><code>alignHorizontal: "left"</code> 和 <code>indentLeft: \<value\></code></li></ul><ul><li><code>alignHorizontal: "right"</code> 和 <code>indentRight: \<value\></code></li></ul><ul><li><code>alignHorizontal: "distributed"</code> 和 <code>indentDistributed: \<value\></code></li></ul>|
| `alignVertical:`|"top" \| "center" \| "bottom" \| "justify" \| "distributed"||



**背景**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `backgroundColor:`|“无”\| \<所有预定义的颜色名称\> \| #RRGGBB|预定义的颜色名称：<br/><br/>"black", "blue", "gray", "green", "orange", "pink", "purple", "red", "teal", "turquoise", "violet", "white", "yellow"|



**边框**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `borderStyle:`|“无”\| \<所有预定义的边框样式名称\>|预定义的边框样式名称：<br/><br/>“dash dot”，“dash dot dot”，“dashed”，“dotted”，“double”，“hair”，“medium dash dot”，“medium dash dot dot”，“medium dashed”，“medium”，“slant dash dot”，“thick”，“thin”<br/><br/>适用于指定范围中的所有边框。 （相当于同时使用“**单元格格式**”对话框中“**边框**”选项卡上的“**外边框**”和“**内边框**”预设指定边框样式。）<br/><br/> **注意：**Excel 2013 支持呈现所有的 13 种预定义边框样式。 但是，Excel Online 并不支持所有边框样式。 下表介绍了打开 Excel Online 中的电子表格时每个边框样式使用的呈现方式。<br/><br/><table><tr><th>Excel 2013</th><th>Excel Online</th></tr><tr><td>“点划线”</td><td>虚线 (1px)</td></tr><tr><td>“短线-点-点”</td><td>点线 (1px)</td></tr><tr><td>“虚线”</td><td>点线 (1px)</td></tr><tr><td>“点线”</td><td>虚线 (1px)</td></tr><tr><td>“双线”</td><td>双线 (3px)</td></tr><tr><td>“极细线”</td><td>实线 (1px)</td></tr><tr><td>“中划线-点”</td><td>虚线 (2px)</td></tr><tr><td>“中划线-点-点”</td><td>点线 (2px)</td></tr><tr><td>“中划线”</td><td>虚线 (2px)</td></tr><tr><td>“中等”</td><td>实线 (2px)</td></tr><tr><td>“斜划线-点”</td><td>虚线 (2px)</td></tr><tr><td>“粗线”</td><td>实线 (3px)</td></tr><tr><td>“细线”</td><td>实线 (1px)</td></tr></table>|
| `borderColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|适用于指定范围中的所有边框。|
| `borderTopStyle:`|“无”\| \<所有预定义的边框样式名称\>|适用于指定范围中的所有边框。|
| `borderTopColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|适用于指定范围中的所有边框。|
| `borderBottomStyle:`|“无”\| \<所有预定义的边框样式名称\>|适用于指定范围中的所有边框。|
| `borderBottomColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|适用于指定范围中的所有边框。|
| `borderLeftStyle:`|“无”\| \<所有预定义的边框样式名称\>|适用于指定范围中的所有边框。|
| `borderLeftColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|适用于指定范围中的所有边框。|
| `borderRightStyle:`|“无”\| \<所有预定义的边框样式名称\>|适用于指定范围中的所有边框。|
| `borderRightColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|适用于指定范围中的所有边框。|
| `borderOutlineStyle:`|“无”\| \<所有预定义的边框样式名称\>|适用于指定范围中的所有边框。|
| `borderOutlineColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|适用于指定范围中的所有边框。|
| `borderInlineStyle:`|“无”\| \<所有预定义的边框样式名称\>|仅适用于指定范围中的内边框。 （相当于仅使用“**单元格格式**”对话框中“**边框**”选项卡上的“**内边框**”预设指定边框样式。）|
| `borderInlineColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB|仅适用于指定范围中的内边框 |



**单元格宽度、高度和换行**


|**密钥**|**值**|
|:-----|:-----|
| `width:`|“自动调整”\|  **数字**|
| `height:`|“自动调整”\|  **数字**|
| `wrapping:`|**布尔**|



**字体**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `fontFamily:`|\<所有可用的字体名称\>|在 Excel Online 中设置字体时，如果字体在浏览器中无法显示，API 将尝试按顺序退回以下字体：Segoe UI、Thonburi、Arial、Verdana 和 Microsoft Sans Serif 字体。如果这些字体均无法显示，则使用浏览器的默认字体。|
| `fontStyle:`|“regular”\| "italic" \| "bold" \| “bold italic”|**注意**：此次发布时，将 `fontStyle:` 设置为“italic”，然后设置为“bold”（或相反）表现为这两种设置的结合。 即，例如如果你先设置“italic”，然后设置“bold”，结果将为“bold italic”。 要_仅_在之前设置为 bold 或 italic 的范围上设置 italic 或 bold，你必须先设置 `fontStyle:"regular"` 以清除之前的格式。|
| `fontSize:`|**数字**||
| `fontUnderlineStyle:`|“无”\| "single" \| "double" \| "single accounting" \| “double accounting”||
| `fontColor:`|“自动”\| \<所有预定义的颜色名称\> \| #RRGGBB||
| `fontDirection:`|“context”\| "left-to-right" \| “right-to-left”|Excel Online 当前不支持从右向左显示文本。 然而，如果你的外接程序在 Excel Online 中运行时，将 `fontDirection:` 设置为“从右向左”，格式设置就保存在工作簿文件中，当在桌面 Desktop Excel 中打开工作簿文件时，即可正确显示。|
| `fontStrikethrough:`|**布尔**||
| `fontSuperscript:`|**布尔**||
| `fontSubScript:`|**布尔**||
| `fontNormal:`|**布尔**|将字体、字体样式、大小和效果设置为正常样式。 这会将单元格字体格式化重置为默认值。 相当于选中“**单元格格式**”对话框“**字体**”选项卡的“**正常字体**”复选框。|



**缩进**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `indentLeft:`|**数字**|与对齐值结合使用时，仅支持以下组合：<br/><br/><ul><li><code>alignHorizontal: "left"</code> 和 <code>indentLeft: \<value\></code></li></ul>|
| `indentRight:`|**数字**|与对齐值结合使用时，仅支持以下组合：<br/><br/><ul><li><code>alignHorizontal: "right"</code> 和 <code>indentRight: \<value\></code></li></ul>|
| `indentDistributed:`|**数字**|与对齐值结合使用时，仅支持以下组合：<br/><br/><ul><li><code>alignHorizontal: "distributed"</code> 和 <code>indentDistributed: \<value\></code></li></ul>|



**数值格式**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `numberFormat:`|**字符串**|要指定数字格式化，请使用自定义数字格式字符串。例如，要指定使用逗号作为千位分隔符的两位小数，您可以指定：<br/><br/> `numberFormat:"#,###.00"`<br/><br/>这些是你可以 [使用“单元格格式”对话框中的“数字”选项卡上的“自定义格式”类别创建](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1) 的相同自定义格式字符串。<br/><br/> **提示：**可以通过以下步骤在 Excel 的“**单元格格式**”对话框查看标准类别的格式字符串：<br/><br/><ol><li>选择一个标准的格式类别，例如“<b>类别</b>”列表中的“<span class="ui">货币</span>。</li><li>设置对话框右侧的格式选项。</li><li>选择“<b>自定义</b>”类别查看“<b>类型</b>”列表的格式字符串。</li></ol>|



**表选项**


|**密钥**|**值**|**注释**|
|:-----|:-----|:-----|
| `style:`|“无”\| \<所有预定义的表格样式名称\>|预定义的表格样式名称：<br/><br/>“TableStyleLight1”、“TableStyleLight2”、“TableStyleLight3”、“TableStyleLight4”、“TableStyleLight5”、“TableStyleLight6”、
“TableStyleLight7”、“TableStyleLight8”、“TableStyleLight9”、“TableStyleLight10”、“TableStyleLight11”、“TableStyleLight12”、“TableStyleLight13”、“TableStyleLight14”、“TableStyleLight15”、“TableStyleLight16”、“TableStyleLight17”、
“TableStyleLight18”、“TableStyleLight19”、“TableStyleLight20”、“TableStyleLight21”、“TableStyleMedium1”、“TableStyleMedium2”、“TableStyleMedium3”、“TableStyleMedium4”、“TableStyleMedium5”、“TableStyleMedium6”、
“TableStyleMedium7”、“TableStyleMedium8”、“TableStyleMedium9”、“TableStyleMedium10”、“TableStyleMedium11”、“TableStyleMedium12”、“TableStyleMedium13”、“TableStyleMedium14”、“TableStyleMedium15”、“TableStyleMedium16”、
“TableStyleMedium17”、“TableStyleMedium18”、“TableStyleMedium19”、“TableStyleMedium20”、“TableStyleMedium21”、“TableStyleMedium22”、“TableStyleMedium23”、“TableStyleMedium24”、“TableStyleMedium25”、“TableStyleMedium26”、
“TableStyleMedium27”、“TableStyleMedium28”、“TableStyleDark1”、“TableStyleDark2”、“TableStyleDark3”、“TableStyleDark4”、“TableStyleDark5”、“TableStyleDark6”、“TableStyleDark7”、“TableStyleDark8”、“TableStyleDark9”、
“TableStyleDark10”、“TableStyleDark11”<br/><br/>若要查看表格样式，请在 Excel 中插入表格，在“**表格工具**” \| “**设计**”选项卡上，选择“**快速样式**”下拉列表，然后选择一个预定义的样式。 此样式的工具提示将与上述列表中值之一相对应。|
| `headerRow:`|**布尔**||
| `firstColumn:`|**布尔**||
| `filterButton:`|**布尔**||
| `totalRow:`|**布尔**||
| `lastColumn:`|**布尔**||
| `bandedRows:`|**布尔**||
| `bandedColumns:`|**布尔**||
