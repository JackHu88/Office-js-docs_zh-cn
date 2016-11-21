# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。有关详细信息，请参阅[指定 Office 主机和 API 要求](../docs/overview/specify-office-hosts-and-api-requirements.md)。

Excel 加载项在多个 Office 版本中运行，包括 Office 2016 for Windows、Office for iPad、Office for Mac 和 Office Online。下表列出了 Excel 要求集、支持该要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。 

|  要求集  |  Office 2016 for Windows*  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |
|:-----|-----|:-----|:-----|:-----|
| ExcelApi 1.3  | 版本 1608 或更高版本| 1.27 或更高版本 |  15.27 或更高版本| 2016 年 9 月 | 
| ExcelApi 1.2  | 版本 1601 或更高版本 | 1.21 或更高版本 | 15.22 或更高版本| 2016 年 1 月 |
| ExcelApi 1.1  | 版本 1509（内部版本 4266.1001）或更高版本 | 1.19 或更高版本 | 15.20 或更高版本| 2016 年 1 月 |

> &#42;**注意**：通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1 要求集。

若要详细了解版本号和内部版本号，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [使用的是哪一版 Office？](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集
若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新 
下面介绍了要求集 1.3 中 Excel JavaScript API 的新增内容。 

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[binding](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/binding.md)|_方法_ > [delete()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/binding.md#delete)|删除 binding 对象。|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_方法_ > [add(range:区域或字符串, bindingType: 字符串, id: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|将新的 binding 对象添加到特定区域。|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_方法_ > [addFromNamedItem(name: 字符串, bindingType: 字符串, id: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|根据工作簿中的命名项添加新的 binding 对象。|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_方法_ > [addFromSelection(bindingType: 字符串, id: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|根据当前选择的内容添加新的 binding 对象。|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_方法_ > [getItemOrNull(id: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#getitemornullid-string)|按 ID 获取 binding 对象。如果 binding 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[chartCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/chartcollection.md)|_方法_ > [getItemOrNull(name: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/chartcollection.md#getitemornullname-string)|按图表名称获取 chart 对象。如果有多个同名的 chart 对象，则此方法返回第一个对象。|1.3|
|[namedItemCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/nameditemcollection.md)|_方法_ > [getItemOrNull(name: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/nameditemcollection.md#getitemornullname-string)|按 nameditem 对象的名称获取此对象。如果 nameditem 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_属性_ > name|PivotTable 对象的名称。|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_关系_ > worksheet|包含当前 PivotTable 对象的工作表。只读。|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_方法_ > [refresh()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md#refresh)|刷新 PivotTable 对象。|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_属性_ > items|一组 PivotTable 对象。只读。|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_方法_ > [getItem(name: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md#getitemname-string)|按名称获取 PivotTable 对象。|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_方法_ > [getItemOrNull(name: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md#getitemornullname-string)|按名称获取 PivotTable 对象。如果 PivotTable 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[range](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md)|_方法_ > [getIntersectionOrNull(anotherRange:区域或字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md#getintersectionornullanotherrange-range-or-string)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.3|
|[range](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md)|_方法_ > [getVisibleView()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md#getvisibleview)|表示当前 range 对象的可见行。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > cellAddresses|表示 RangeView 的单元格地址。只读。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > columnCount|返回可见列数。只读。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > formulas|表示采用 A1 表示法的公式。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > formulasLocal|使用用户语言和数字格式区域设置表示采用 A1 表示法的公式。例如，用英语表示的公式 "=SUM(A1, introduced in 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > index|返回表示 RangeView 的索引的值。只读。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > numberFormat|表示 Excel 中指定单元格的数字格式代码。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > rowCount|返回可见行数。只读。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > text|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > valueTypes|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_属性_ > values|表示指定的 RangeView 的原始值。返回的数据可能是字符串、数字，也可能是布尔值。包含错误的单元格将返回错误字符串。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_关系_ > rows|表示一组与 range 相关联的 RangeView。只读。|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_方法_ > [getRange()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md#getrange)|获取与当前 RangeView 相关联的父 range。|1.3|
|[rangeViewCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md)|_属性_ > items|一组 rangeView 对象。只读。|1.3|
|[rangeViewCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md)|_方法_ > [getItemAt(index: 数字)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md#getitematindex-number)|按索引获取 RangeView 行。从零开始编制索引。|1.3|
|[setting](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.3|
|[setting](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md)|_方法_ > [delete()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md#delete)|删除 setting 对象。|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_属性_ > items|一组 setting 对象。只读。|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_方法_ > [getItem(key: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#getitemkey-string)|按键获取 setting 项。|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_方法_ > [getItemOrNull(key: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#getitemornullkey-string)|按键获取 setting 项。如果 setting 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_方法_ > [set(key: 字符串, value: 字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#setkey-string-value-string)|设置指定的 setting 对象，或将其添加到工作簿中。|1.3|
|[settingsChangedEventArgs](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingschangedeventargs.md)|_关系_ > settingCollection|获取表示引发了 SettingsChanged 事件的 binding 的 setting 对象。|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_属性_ > highlightFirstColumn|指明第一列是否包含特殊格式。|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_属性_ > highlightLastColumn|指明最后一列是否包含特殊格式。|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_属性_ > showBandedColumns|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_属性_ > showBandedRows|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_属性_ > showFilterButton|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|1.3|
|[tableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecollection.md)|_方法_ > [getItemOrNull(key: 数字或字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecollection.md#getitemornullkey-number-or-string)|按名称或 ID 获取 table 对象。如果 table 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[tableColumnCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecolumncollection.md)|_方法_ > [getItemOrNull(key: 数字或字符串)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecolumncollection.md#getitemornullkey-number-or-string)|按名称或 ID 获取 column 对象。如果 column 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[workbook](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/workbook.md)|_关系_ > pivotTables|表示一组与 workbook 相关联的 PivotTable 对象。只读。|1.3|
|[workbook](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/workbook.md)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.3|
|[worksheet](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/worksheet.md)|_关系_ > pivotTables|一组属于 worksheet 的 PivotTable 对象。只读。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 的最近更新
下面介绍了要求集 1.2 中 Excel JavaScript API 的新增内容。 

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_属性_ > id|按 chart 对象在集合中的位置获取此对象。只读。|1.2|
|[chart](../excel/chart.md)|_关系_ > worksheet|包含当前 chart 的 worksheet 对象。只读。|1.2|
|[chart](../excel/chart.md)|_方法_ > [getImage(height: 数字, width: 数字, fittingMode: 字符串)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|通过缩放 chart 以适应指定的尺寸，将 chart 呈现为 base64 编码的图像。|1.2|
|[filter](../excel/filter.md)|_关系_ > criteria|当前在给定列上应用的 filter。只读。|1.2|
|[filter](../excel/filter.md)|_方法_ > [apply(criteria:FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|在给定列中应用给定的筛选条件。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyBottomItemsFilter(count: 数字)](../excel/filter.md#applybottomitemsfiltercount-number)|将“Bottom Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyBottomPercentFilter(percent: 数字)](../excel/filter.md#applybottompercentfilterpercent-number)|将“Bottom Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyCellColorFilter(color: 字符串)](../excel/filter.md#applycellcolorfiltercolor-string)|将“Cell Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyCustomFilter(criteria1: 字符串, criteria2: 字符串, oper: 字符串)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|将“Icon”筛选器应用于列，以获取给定的条件字符串。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyDynamicFilter(criteria: 字符串)](../excel/filter.md#applydynamicfiltercriteria-string)|将“Dynamic”筛选器应用于列。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyFontColorFilter(color: 字符串)](../excel/filter.md#applyfontcolorfiltercolor-string)|将“Font Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyIconFilter(icon:Icon)](../excel/filter.md#applyiconfiltericon-icon)|将“Icon”筛选器应用于列，以获取给定 icon。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyTopItemsFilter(count: 数字)](../excel/filter.md#applytopitemsfiltercount-number)|将“Top Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyTopPercentFilter(percent: 数字)](../excel/filter.md#applytoppercentfilterpercent-number)|将“Top Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](../excel/filter.md)|_方法_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|将“Values”筛选器应用于列，以获取给定值。|1.2|
|[filter](../excel/filter.md)|_方法_ > [clear()](../excel/filter.md#clear)|清除给定列上的 filter。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > color|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选器结合使用。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > criterion1|第一个用于筛选数据的条件。在“custom”筛选器中用作运算符。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > criterion2|第二个用于筛选数据的条件。在“custom”筛选器中仅用作运算符。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > dynamicCriteria|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > filterOn|filter 使用的属性，用于确定值是否应一直可见。可取值为：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > operator|使用“自定义”筛选器时，用于组合条件 1 和 2 的运算符。可取值为：And、Or。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_属性_ > values|一组用于“values”筛选器的值。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_关系_ > icon|用于筛选单元格的 icon。与“icon”筛选器结合使用。|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_属性_ > date|用于筛选数据的采用 ISO8601 格式的日期。|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_属性_ > specificity|用于保留数据的日期的具体程度。例如，如果日期是 2005-04-02 并且将特殊性设置为“月”，则筛选操作将保留包含 2009 年 4 月日期的所有行。可能的值是：Year、Monday、Day、Hour、Minute、Second。|1.2|
|[formatProtection](../excel/formatprotection.md)|_属性_ > formulaHidden|表示 Excel 是否隐藏区域中的单元格公式。指示整个区域不具有统一公式隐藏设置的空值。|1.2|
|[formatProtection](../excel/formatprotection.md)|_属性_ > locked|指明 Excel 是否锁定对象中的单元格。若值为 Null，则表明整个 range 的锁定设置不一致。|1.2|
|[icon](../excel/icon.md)|_属性_ > index|表示 icon 在给定集中的索引。|1.2|
|[icon](../excel/icon.md)|_属性_ > set|表示图标所属的集合。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[range](../excel/range.md)|_属性_ > columnHidden|表示当前 range 的所有列均已隐藏。|1.2|
|[range](../excel/range.md)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.2|
|[range](../excel/range.md)|_属性_ > hidden|表示当前 range 的所有单元格均已隐藏。只读。|1.2|
|[range](../excel/range.md)|_属性_ > rowHidden|表示当前 range 的所有行均已隐藏。|1.2|
|[range](../excel/range.md)|_关系_ > sort|表示当前 range 的区域排序。只读。|1.2|
|[range](../excel/range.md)|_方法_ > [merge(across: 布尔值)](../excel/range.md#mergeacross-bool)|将 range 单元格合并到 worksheet 的一个区域内。|1.2|
|[range](../excel/range.md)|_方法_ > [unmerge()](../excel/range.md#unmerge)|将 range 单元格拆分为单个单元格。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_属性_ > columnWidth|获取或设置 range 内所有列的宽度。如果列宽不一致，则此方法返回 NULL。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_属性_ > rowHeight|获取或设置 range 内所有行的高度。如果行高不一致，则此方法返回 NULL。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_关系_ > protection|返回 range 的格式保护对象。只读。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_方法_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|根据列中的当前数据，更改当前 range 内所有列的宽度，以达到最佳显示效果。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_方法_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|根据列中的当前数据，更改当前 range 内所有行的高度，以达到最佳显示效果。|1.2|
|[rangeReference](../excel/rangereference.md)|_属性_ > address|表示当前 range 对象的可见行。|1.2|
|[rangeSort](../excel/rangesort.md)|_方法_ > [apply(fields:SortField[], matchCase: 布尔值, hasHeaders: 布尔值, orientation: 字符串, method: 字符串)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|执行排序操作。|1.2|
|[sortField](../excel/sortfield.md)|_属性_ > ascending|表示是否执行升序排序。|1.2|
|[sortField](../excel/sortfield.md)|_属性_ > color|表示按字体或单元格颜色进行排序时，条件的目标颜色。|1.2|
|[sortField](../excel/sortfield.md)|_属性_ > dataOption|表示此字段的其他排序选项。可能的值是：Normal、TextAsNumber。|1.2|
|[sortField](../excel/sortfield.md)|_属性_ > key|表示条件所在的列（或行，具体取决于排序方向）。表示与第一列（或行）的偏移量。|1.2|
|[sortField](../excel/sortfield.md)|_属性_ > sortOn|表示此排序条件的类型。可能的值是：Value、CellColor、FontColor、Icon。|1.2|
|[sortField](../excel/sortfield.md)|_关系_ > icon|表示对单元格图标进行排序时，条件的目标图标。|1.2|
|[table](../excel/table.md)|_关系_ > sort|表示表的排序。只读。|1.2|
|[table](../excel/table.md)|_关系_ > worksheet|包含当前 table 的 worksheet 对象。只读。|1.2|
|[table](../excel/table.md)|_方法_ > [clearFilters()](../excel/table.md#clearfilters)|清除当前在 table 中应用的所有 filter。|1.2|
|[table](../excel/table.md)|_方法_ > [convertToRange()](../excel/table.md#converttorange)|将 table 转换为包含单元格的普通 range。保留所有数据。|1.2|
|[table](../excel/table.md)|_方法_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|重新应用当前在 table 上应用的所有 filter。|1.2|
|[tableColumn](../excel/tablecolumn.md)|_关系_ > filter|检索应用于列的 filter。只读。|1.2|
|[tableSort](../excel/tablesort.md)|_属性_ > matchCase|表示最后一次对表进行排序时大小写是否有影响。只读。|1.2|
|[tableSort](../excel/tablesort.md)|_属性_ > method|表示最后一次对表进行排序时所使用的中文字符排序方法。只读。可能的值是：PinYin、StrokeCount。|1.2|
|[tableSort](../excel/tablesort.md)|_关系_ > fields|表示最后一次对表进行排序时所使用的当前条件。只读。|1.2|
|[tableSort](../excel/tablesort.md)|_方法_ > [apply(fields:SortField[], matchCase: 布尔值, method: 字符串)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|执行排序操作。|1.2|
|[tableSort](../excel/tablesort.md)|_方法_ > [clear()](../excel/tablesort.md#clear)|清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。|1.2|
|[tableSort](../excel/tablesort.md)|_方法_ > [reapply()](../excel/tablesort.md#reapply)|对 table 重新应用当前的排序参数。|1.2|
|[workbook](../excel/workbook.md)|_关系_ > functions|表示包含此 workbook 的 Excel 应用程序实例。只读。|1.2|
|[worksheet](../excel/worksheet.md)|_关系_ > protection|返回 worksheet 的工作表保护对象。只读。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_属性_ > protected|指明 worksheet 是否受保护。只读。只读。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_关系_ > options|工作表保护选项。只读。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_方法_ > [protect(options:WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|保护 worksheet。如果 worksheet 处于受保护状态，则无法执行此方法。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_方法_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|解除对 worksheet 的保护。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowAutoFilter|表示允许使用自动筛选功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowDeleteColumns|表示允许删除列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowDeleteRows|表示允许删除行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowFormatCells|表示允许格式化单元格的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowFormatColumns|表示允许格式化列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowFormatRows|表示允许格式化行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowInsertColumns|表示允许插入列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowInsertHyperlinks|表示允许插入超链接的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowInsertRows|表示允许插入行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowPivotTables|表示允许使用数据透视表功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_属性_ > allowSort|表示允许使用排序功能的工作表保护选项。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1
Excel JavaScript API 1.1 是首版 API。有关 API 的详细信息，请参阅“Excel JavaScript API”参考主题。  
    
## <a name="additional-resources"></a>其他资源

- [指定 Office 主机和 API 要求](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
