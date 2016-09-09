# GetStarted 元素

提供在 Word、Excel、PowerPoint 和 OneNote 主机中安装此外接程序时显示的标注所使用的信息。 **GetStarted** 元素是 [FormFactor](./formfactor.md) 的子元素。

## 子元素

| 元素                       | 必需 | 说明                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [标题](#标题)               | 是      | 定义外接程序公开功能的位置。     |
| [说明](#说明)   | 是      | 包含 JavaScript 函数的文件的 URL。|
| [LearnMoreUrl](#learnmoreurl) | 否       | 指向详细说明外接程序的页面的 URL。   |


## 标题 
必需。 用于标注顶部的标题。 **resid** 属性引用 [Resources](./resources.md) 分区的 [ShortStrings](./resources.md#shortstrings) 元素中的有效 ID。

## 说明
必需。 标注的说明/正文内容。 **resid** 属性引用 [Resources](./resources.md) 分区的 [LongStrings](./resources.md#longstrings) 元素中的有效 ID。

## LearnMoreUrl
必需。 指向用户可以了解你的外接程序详细信息的页面 URL。 **resid** 属性引用 [Resources](./resources.md) 分区的 [Urls](./resources.md#urls) 元素中的有效 ID。

> **注意：****LearnMoreUrl** 当前无法在 Word、Excel 或 PowerPoint 客户端中呈现。 我们建议为所有客户端添加此 URL，以便 URL 在可用时呈现。 
