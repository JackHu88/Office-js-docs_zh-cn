# Excel 外接程序 JavaScript API 参考

_适用于：Excel 2016、Office 2016_

下面的链接显示了 API 中可用的高级 Excel 对象。每个对象页面链接包含对象可用的属性、关系和方法的描述。如需了解详细信息，请浏览下面的链接。
	
* [工作簿](resources/workbook.md)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。 
* [工作表](resources/worksheet.md)：工作表集合的成员。工作表集合包含工作簿中的所有 Workbook 对象。
	* [工作表集合](resources/worksheetcollection.md)：属于工作簿的所有 Workbook 对象的集合。 
* [区域](resources/range.md)：表示某一单元格、某一行、某一列、某一选定区域（该区域可包含一个或若干连续单元格区域）。  
* [表](resources/table.md)：表示有组织的单元格的集合，设计用于简化数据管理。 
	* [表集合](resources/tablecollection.md)：工作簿或工作表中的表的集合。 
	* [TableColumn 集合](resources/tablecolumncollection.md)：表中所有列的集合。 
	* [TableRow 集合](resources/tablerowcollection.md)：表中所有行的集合。 
* [图表](resources/chart.md)：表示工作表中的 Chart 对象，它是基础数据的可视表示形式。   
	* [图表集合](resources/chartcollection.md)：工作表中的图表的集合。	
* [NamedItem](resources/nameditem.md)：表示单元格区域或值的定义名称。名称可以是基元命名的对象、range 对象等。
	* [NamedItem 集合](resources/nameditemcollection.md)：工作簿中 NamedItem 对象的集合。
* [绑定](resources/binding.md)：表示对工作簿的某一部分的绑定的抽象类。
	* [绑定集合](resources/bindingcollection.md)：表示属于工作簿的所有绑定对象的集合。 
* [TrackedObject 集合](resources/trackedobjectscollection.md)：允许外接程序管理各 sync() 批次的 range 对象引用。 
* [请求上下文](resources/requestcontext.md)：RequestContext 对象可加快对 Excel 应用程序的请求。


##### 其他资源

*  [Excel 外接程序编程概述](excel-add-ins-programming-overview.md)
*  [构建第一个 Excel 外接程序](build-your-first-excel-add-in.md)
*  [适用于 Excel 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Excel 外接程序代码示例](excel-add-ins-code-samples.md) 


