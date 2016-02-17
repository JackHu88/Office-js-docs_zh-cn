# Word 外接程序 JavaScript 参考 

对于 Word 外接程序，查找适用于 Word 的 JavaScript API 的 API 参考。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 本节内容

这些是 Word JavaScript API 的主要对象。

* [Body](word-add-ins-javascript-reference/body.md)：表示文档或节的正文。
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md)：内容的容器。它是文档中可能标记的绑定区域，可作为特定类型的内容的容器。例如，内容控件可能包含格式化文本段落等内容及其他内容控件。您可以通过文档、文档正文、段落和区域的内容控件集合，或者在内容控件上访问内容控件。
* [Document](word-add-ins-javascript-reference/document.md)：顶层的对象。Document 对象包含一个或多个[节](word-add-ins-javascript-reference/section.md)，它是包含文档内容以及标头/页脚信息的正文。
* [Font](word-add-ins-javascript-reference/font.md)：为正文、内容控件、段落或区域提供文本格式设置。
* [Image](word-add-ins-javascript-reference/inlinepicture.md)：表示固定到段落的嵌入式图片。
* [Paragraph](word-add-ins-javascript-reference/paragraph.md)：表示选定内容、区域或文档中的单个段落。您可以通过选定内容、区域或文档中的段落集合访问段落。 
* [Range](word-add-ins-javascript-reference/range.md)：表示文档中的相邻区域。当您获取选定内容、将内容插入正文、将内容插入内容控件、将内容插入段落或者获取搜索结果时，您将获得一个 Range 对象。您可以定义区域并对其进行操作，无需更改选定内容。
* [Section](word-add-ins-javascript-reference/section.md)：定义文档的不同标头和页脚以及不同的页面布局配置。您可以从 Document 对象访问节。 
* [Selection](word-add-ins-javascript-reference/document.md#getselection)：Document 对象允许您访问用户在文档中的选定内容或当前插入点（如果未选定任何内容）。

## 向我们提供反馈

您的反馈对我们意义重大。 

* 查看文档并在此存储库中直接[提交问题](https://github.com/OfficeDev/office-js-docs/issues)，告诉我们您在其中发现的任何疑问和问题。
* 让我们了解您的编程体验、您希望在未来版本中看到的功能、代码示例，等等。请在[此网站](http://officespdev.uservoice.com/)输入您的建议和想法。

## 其他资源

* [Word 外接程序](word-add-ins.md)
* [Word 外接程序编程指南](word-add-ins-programming-guide.md)
* [Office 外接程序](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [开始使用 Office 外接程序](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;GitHub 上的 Word 外接程序&lt;/a&gt;
* [适用于 Word 的代码段资源管理器](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
