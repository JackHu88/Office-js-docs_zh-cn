
# <a name="office-addin-requirement-sets"></a>Office 外接程序要求集

要求集是 API 成员的命名组。Office 外接程序使用清单中指定的要求集或使用运行时检查以确定 Office 主机是否支持外接程序所需的 API。有关详细信息，请参阅[指定 Office 主机和 API 要求](../docs/overview/specify-office-hosts-and-api-requirements.md)。

若要更广泛地获取 Office 主机支持的外接程序的平台的信息，请查看 [Office 外接程序主机和平台可用性](https://dev.office.com/add-in-availability)页。

## <a name="requirement-sets"></a>要求集


下表列出了要求集的名称、每个集中的方法、支持该要求集的 Office 主机应用程序以及 API 的版本号。

有关 Outlook 的要求集的信息，请参阅 [了解 Outlook API 要求集](./outlook/tutorial-api-requirement-sets.md)

|  集名称  |  版本  |  Office 主机  |  集中的方法  |
|:-----|-----|:-----|:-----|
| ExcelApi   | 1.2 | Excel 2016<br>Excel Online<br>Excel for iPad<br>|工作表保护<br>工作表函数<br>排序<br>筛选<br>R1C1 参考样式<br>合并单元格<br>调整行高和列宽<br>Chart.getImage()<br>Range.getUsedRange(valuesOnly)|
| ExcelApi   | 1.1 | Excel 2016<br>Excel Online<br>Excel for iPad<br>|Excel 命名空间中的所有元素|
| WordApi    | 1.2 | Word 2016<br>Word 2016 for Mac<br>Word for iPad<br>Word Online| Word 命名空间中的所有元素。以下方法被添加到了该 WordApi 版本：<br>Body.select(selectionMode)<br>Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>contentControl.select(selectionMode)<br>contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>inlinePicture.paragraph<br>inlinePicture.delete<br>inlinePicture.insertBreak(breakType, insertLocation)<br>inlinePicture.insertFileFromBase64(base64file, insertLocation)<br>inlinePicture.insertHtml(html, insertLocation)<br>inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)<br>inlinePicture.insertOoxml(ooxml, insertLocation)<br>inlinePicture.insertParagraph(paragraphText, insertLocation)<br>inlinePicture.insertText(text, insertLocation)<br>inlinePicture.select(selectionMode)<br>paragraph.select(selectionMode)<br>range.inlinePictures<br>range.select(selectionMode)<br>range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|
| WordApi    | 1.1 | Word 2016<br>Word 2016 for Mac<br>Word for iPad<br>Word Online|Word 命名空间中的所有元素（已添加到 WordApi 1.2 及更高版本的 API 成员除外），如上所列。|
| ActiveView | 1.1 | PowerPoint<br>PowerPoint Online|Document.getActiveViewAsync|
| BindingEvents  | 1.1 | Access Web App<br>Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | 1.1 |PowerPoint<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|支持使用 Document.getFileAsync 方法时输出作为字节数组 (Office.FileType.Compressed) 的 Office Open XML (OOXML) 格式<br>。|
| CustomXmlParts    | 1.1 |Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogAPI | 1.1 | Excel<br>PowerPoint<br>Word 2016<br>Outlook|Office.context.ui.displayDialogAsync()<br>Office.context.ui.messageParent()<br>Office.context.ui.close()|
| DocumentEvents    | 1.1 | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| 文件  | 1.1 | PowerPoint<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | 1.1 | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 HTML (Office.CoercionType.Html)<br>。|
| ImageCoercion | 1.1 | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.setSelectedDataAsync 方法写入数据时转换为图像 (Office.CoercionType.Image)。|
| 邮箱   |   | Outlook for Windows<br>Outlook for web<br>Outlook for Mac<br>Outlook Web App |查看 [了解 Outlook API 要求集](./outlook/tutorial-api-requirement-sets.md)|
| MatrixBindings    | 1.1 | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | 1.1 | Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“矩阵”（数组的数组）数据结构 (Office.CoercionType.Matrix)。|
| OoxmlCoercion | 1.1 | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 Open Office XML (OOXML) 格式 (Office.CoercionType.Ooxml)。|
| PartialTableBindings  | 1.1 | Access Web App||
| PdfFile   | 1.1 | PowerPoint<br/>PowerPoint Online<br/>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持使用 Document.getFileAsync 方法时输出 PDF 格式 (Office.FileType.Pdf)<br>。|
| 选择 | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>项目<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| 设置  | 1.1 | Access Web App<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | 1.1 | Access Web App<br>Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | 1.1 | Access Web App<br>Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“表格”数据结构 (Office.CoercionType.Table)。|
| TextBindings  | 1.1 | Excel<br>Excel Online<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>项目<br>Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为文本格式 (Office.CoercionType.Text)。|
| TextFile  | 1.1 | Word 2013 及更高版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|支持在使用 Document.getFileAsync 方法时输出文本格式 (Office.FileType.Text)。|

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

