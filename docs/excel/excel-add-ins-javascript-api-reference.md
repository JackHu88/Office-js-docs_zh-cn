# <a name="excel-javascript-api-reference"></a>Excel JavaScript API 参考

你可以使用 Excel JavaScript API 构建适用于 Excel 2016 的外接程序。以下列表显示在 API 中可用的高级 Excel 对象。每个对象页面链接包含对象可用的属性、关系和方法的描述。如需了解详细信息，请浏览相应链接。

* [工作簿](../../reference/excel/workbook.md)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。
* [Worksheet](../../reference/excel/worksheet.md)：工作表集合的成员。工作表集合包含工作簿中的所有 Workbook 对象。
    * [WorksheetCollection](../../reference/excel/worksheetcollection.md)：属于工作簿的所有 Worksheet 对象的集合。
* [Range](../../reference/excel/range.md)：表示某一单元格、某一行、某一列、某一选定区域（其中包含一个或多个相邻单元格块）。
* [表](../../reference/excel/table.md)：表示有组织的单元格的集合，设计用于简化数据管理。
    * [表集合](../../reference/excel/tablecollection.md)：工作簿或工作表中的表的集合。
    * [TableColumn 集合](../../reference/excel/tablecolumncollection.md)：表中所有列的集合。
    * [TableRow 集合](../../reference/excel/tablerowcollection.md)：表中所有行的集合。
* [图表](../../reference/excel/chart.md)：表示工作表中的 Chart 对象，它是基础数据的可视表示形式。
    * [图表集合](../../reference/excel/chartcollection.md)：工作表中的图表的集合。
* [TableSort](../../reference/excel/tablesort.md)：表示对 Table 对象上的操作进行排序的对象。
* [RangeSort](../../reference/excel/rangesort.md)：表示对 Range 对象上的操作进行排序的对象。
* [筛选器](../../reference/excel/filter.md)：表示管理表格列筛选的筛选器对象。
* [工作表保护](../../reference/excel/worksheetprotection.md)：表示对工作表对象的保护。
* [工作表函数](../../reference/excel/functions.md)：表示可从 JavaScript 中调用的 Microsoft Excel 工作表函数的容器。
* [NamedItem](../../reference/excel/nameditem.md)：表示单元格区域或值的定义名称。名称可以是基元命名的对象、range 对象等。
    * [NamedItem 集合](../../reference/excel/nameditemcollection.md)：工作簿中 NamedItem 对象的集合。
* [绑定](../../reference/excel/binding.md)：表示对工作簿的某一部分的绑定的抽象类。
    * [绑定集合](../../reference/excel/bindingcollection.md)：表示属于工作簿的所有绑定对象的集合。
* [TrackedObject 集合](../../reference/excel/trackedobjectscollection.md)：允许外接程序管理各 sync() 批次的 range 对象引用。
* [请求上下文](../../reference/excel/requestcontext.md)：RequestContext 对象可加快对 Excel 应用程序的请求。


##### <a name="additional-resources"></a>其他资源

*  [Excel 外接程序编程概述](excel-add-ins-javascript-programming-overview.md)
*  [构建你的第一个 Excel 外接程序](build-your-first-excel-add-in.md)
*  [适用于 Excel 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

