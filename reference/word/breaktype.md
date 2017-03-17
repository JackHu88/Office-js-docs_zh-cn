# <a name="breaktype-javascript-api-for-word"></a>BreakType（适用于 Word 的 JavaScript API）

指定分页符的窗体。

_适用于：Word 2016、Word for iPad、Word for Mac、Word Online_

下面是 API 上支持的分隔符类型。

| **值**         | **类型** | **说明**     |
|:-----------------|:--------|:----|
|行| | 换行符。 |
|页面| | 插入点处的分页符。|
|sectionNext| | 分节符在下一页。下一个类型将过时。|
|sectionContinuous| | 新节不包含相应分页符。|
|sectionEven| string | 使下一节从下一偶数页开始的分节符。如果分节符落入偶数页，则 Word 将下一奇数页留为空白。|
|sectionOdd| string | 使下一节从下一奇数页开始的分节符。如果分节符落入奇数页，则 Word 将下一偶数页留为空白。|

## <a name="support-details"></a>支持详细信息
在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。
