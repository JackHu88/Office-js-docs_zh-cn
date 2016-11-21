# <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

要求集是指各组已命名的 API 成员。Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。有关详细信息，请参阅[指定 Office 主机和 API 要求](../docs/overview/specify-office-hosts-and-api-requirements.md)。

若要了解 Office 主机何时支持外接程序，请参阅 [Office 外接程序主机和平台可用性](https://dev.office.com/add-in-availability)。

## <a name="hostspecific-api-requirement-sets"></a>视具体主机而定的 API 要求集

若要了解 Excel、Word、OneNote、Outlook 和 Dialog API 要求集，请参阅：

- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)
- [OneNote JavaScript API 要求集](onenote-api-requirement-sets.md)
- [了解 Outlook API 要求集](../outlook/tutorial-api-requirement-sets.md)
[Dialog API 要求集](dialog-api-requirement-sets.md)

## <a name="common-api-requirement-sets"></a>通用 API 要求集

下表列出了通用 API 要求集、每个集内的方法，以及支持相应要求集的 Office 主机应用程序。这些 API 要求集都是第 1.1 版。


|  要求集  |  Office 主机  |  集内的方法  |
|:-----|-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint&nbsp;Online|Document.getActiveViewAsync|
| BindingEvents  | Access Web 应用<br>Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | PowerPoint<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|支持使用 Document.getFileAsync 方法时输出作为字节数组 (Office.FileType.Compressed) 的 Office Open XML (OOXML) 格式<br>。|
| CustomXmlParts    | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DocumentEvents    | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | PowerPoint<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 HTML (Office.CoercionType.Html)<br>。|
| ImageCoercion | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.setSelectedDataAsync 方法写入数据时转换为图像 (Office.CoercionType.Image)。|
| 邮箱   |Outlook for Windows<br>Outlook for web<br>Outlook for Mac<br>Outlook Web App |请参阅[了解 Outlook API 要求集](./outlook/tutorial-api-requirement-sets.md)。|
| MatrixBindings    | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“矩阵”（数组的数组）数据结构 (Office.CoercionType.Matrix)。|
| OoxmlCoercion | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 Open Office XML (OOXML) 格式 (Office.CoercionType.Ooxml)。|
| PartialTableBindings  | Access Web 应用||
| PdfFile   | PowerPoint<br/>PowerPoint Online<br/>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持使用 Document.getFileAsync 方法时输出 PDF 格式 (Office.FileType.Pdf)<br>。|
| Selection | Excel<br>Excel Online<br>PowerPoint<br>项目<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web 应用<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web 应用<br>Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web 应用<br>Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“表格”数据结构 (Office.CoercionType.Table)。|
| TextBindings  | Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>PowerPoint<br>项目<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为文本格式 (Office.CoercionType.Text)。|
| TextFile  | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|支持在使用 Document.getFileAsync 方法时输出文本格式 (Office.FileType.Text)。|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>不作为要求集一部分的方法

适用于 Office 的 JavaScript API 中的以下方法不是要求集的一部分。如果外接程序需要这些方法的任意一个，请使用外接程序清单中的 **Methods** 和 **Method** 元素以声明需要这些方法，或使用 if 语句执行运行时检查。有关详细信息，请参阅 [指定 Office 主机和 API 要求](../docs/overview/specify-office-hosts-and-api-requirements.md)。

|**方法名称**|**Office 主机支持**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access Web App、Excel 和 Excel Online|
|Document.getFilePropertiesAsync|Excel、Excel Online、Word 和 PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint 和 PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.goToByIdAsync|Excel、Excel Online、Word 和 PowerPoint|
|Settings.addHandlerAsync|Access Web App、Excel、Excel Online、Word 和 PowerPoint|
|Settings.refreshAsync|Access Web App、Excel、Excel Online、Word、PowerPoint 和 PowerPoint Online|
|Settings.removeHandlerAsync|Access Web App、Excel、Excel Online、Word 和 PowerPoint|
|TableBinding.clearFormatsAsync|Excel 和 Excel Online|
|TableBinding.setFormatsAsync|Excel 和 Excel Online|
|TableBinding.setTableOptionsAsync|Excel 和 Excel Online|

## <a name="additional-resources"></a>其他资源

- [指定 Office 主机和 API 要求](../docs/overview/specify-office-hosts-and-api-requirements.md)



