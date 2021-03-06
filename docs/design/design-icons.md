# <a name="design-icons-for-add-in-commands"></a>外接程序命令的设计图标

[外接程序命令](add-in-commands.md)添加按钮、文本和 Office UI 图标。外接程序命令按钮应提供有意义的图标和标签，以便清楚地标识用户在使用命令时执行的操作。本文提供了样式和生产准则，可帮助你设计与 Office 无缝集成的图标。 

## <a name="office-icon-design-principles"></a>Office 图标设计原则

Office 桌面客户端的 Office 2013 版本包括刷新的图标。替代样式更改已缩减。新图标仅包括必需通信元素。包括透视、渐变和光源的非必需元素均被删除。简化后的图标可支持对命令和控件的快速解析。请按照此样式设计最适合 Office 的图标。

Office 图标均基于以下设计原则完成： 

- 以现代方式阐释 Office 图标集合 
- 全新设计但又不陌生  
- 简单、清楚和直接 

下图显示了应用现代设计原则的图标。

![显示 Office 旧图标的图像和刷新的以现代方式阐释的图标](../../images/icons_image.PNG)

## <a name="icon-guidelines"></a>图标准则
创建图标时，请遵循以下准则： 

- 采用 1 像素网格并使用位图编辑工具，以获得最佳效果。  
- 重绘，但不重设大小。在将图标重设为更大或更小尺寸时，请花时间重绘切割区、角和圆边，以最大化线条的清晰度。 
- 删除使图标显得杂乱的部分。
- 不在 Office 功能区或关联菜单中重复使用 Office UI Fabric 图标。Fabric 图标风格不同，不能匹配。 
- 避免依赖徽标或品牌传达外接程序命令应起到的作用。品牌标志在较小的图标尺寸上和应用很多修饰符后并非总具有识别性。品牌标志经常与 Office 功能区图标样式冲突，并可能在饱和的环境中过度吸引用户的注意力。
- 为辅助功能使用白色填充。图标中的大部分对象都需使用白色背景，以使其在 Office UI 主题中以及高对比度模式下清晰可辨。  
- 使用具有透明背景的 PNG 格式。 
- 避免在图标中使用可本地化的内容，包括印刷字符、段落标记指示和问号。 
- 不要对不同的命令重复使用视觉隐喻。对不同的操作使用同一图标可能会引起混淆。 
- 使您的按钮标签清晰、简洁。将视觉和文本信息结合使用以传达含义。 


## <a name="icon-size-recommendations-and-requirements"></a>图标大小的建议和要求

Office 2016 桌面图标是位图图像。根据用户的 DPI 设置和触摸模式将呈现不同的大小。包括所有八种支持的大小，可在所有受支持的解决方案和上下文中创建最佳体验。以下是受支持的大小 - 三种是必需的：

- 16 像素（必需）
- 20 像素
- 24 像素
- 32 像素（必需）
- 40 像素
- 48 像素
- 64 像素（建议，最适用于 Mac）
- 80 像素（必需）  

确保根据每个尺寸重新绘制你的图标，而非将其缩小。

![显示调整图标大小而非缩小图标的建议的图示](../../images/icon_resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

>**Note:** At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>图标分析和布局

Office 图标通常是由具有操作和概念修饰符的基本元素构成的。 操作修饰符表示诸如添加、打开、新建或关闭等的概念。概念修饰符表示图标的状态、更改或说明。 

若要创建与 Office UI 相符的命令，请遵循基本元素和修饰符的布局准则。这将确保命令看起来具有专业性，且客户将信任你的外接程序。如果出现未按这些准则进行操作的情况，则这些操作应该是有意为之。

以下图像显示 Office 图标中的基本元素和修饰符的布局。

![显示处于中间位置的图标基本元素的图像，其中修饰符位于右下方，操作修饰符位于左上方](../../images/icon_layout.PNG)

- 将基本元素置于像素框架的中间位置，并在其周围填充空白。
- 将操作修饰符置于左上方。 
- 将概念修饰符置于右下方。
- 限制图标中的元素数。在 32 像素中，将修饰符数限制为最多两个。在 16 像素中，将修饰符数限制为一个。

放置与大小相一致的基本元素。如果基本元素不能在框架居中显示，则将其对齐到左上方，并将多余的像素保留在右下方。为了获得最佳效果，请应用下表中列出的填充准则。

|**图标大小**|**在基本元素周围填充**|
|:---|:---|
|16px|0|
|20px|1px|
|24px|1px|
|32px|2px|
|40px|2px|
|48px|3px|
|64px|5px|
|80px|5px|

所有修饰符在每个元素间都应具有 1 像素的透明切割区，包括背景。元素不应直接重叠。在规则和边缘之间创建空白空间。修饰符在大小上可能略有不同，但会将这些尺寸作为起点使用。

|**图标大小**|**修饰符大小**|
|:---|:---|
|16px|9px|
|20px|10px|
|24px|12px|
|32px|14px|
|40px|20px|
|48px|22px|
|64px|29px|
|80px|38px|

## <a name="icon-colors"></a>图标颜色

Office 图标具有一个有限的调色板。使用下表中列出的颜色确保与 Office UI 无缝集成。对颜色使用应用以下准则： 

- 使用颜色传达图标含义，而非只是用作修饰。图标颜色应突出显示或强调操作、状态或明确区分标记的元素。  
- 如有可能，除灰色外仅使用其他一种颜色。将附加颜色限制为最多两种。
- 所有图标大小中的颜色应具有一致的外观。Office 图标针对不同的图标大小具有略微不同的调色板。16 像素和更小的图标稍暗，而与 32 像素和更大的图标相比更亮。除了这些细微的调整以外，颜色的差别体现在大小上。   

|**颜色名称**|**RGB**|**十六进制**|**颜色**|**类别**|
|:---|:---|:---|:---|:---|
|文本灰色 (80)|80、80、80|#505050|![文本灰色 80 彩色图像](../../images/textGray_80.gif)|文本|
|文本灰色 (95)|95、95、95|#5F5F5F|![文本灰色 95 彩色图像](../../images/textGray_95.gif)|文本|
|文本灰色 (105)|105、105、105|#696969|![文本灰色 105 彩色图像](../../images/textGray_105.gif)|文本|
|深灰色 32|128、128、128|#808080|![深灰色 32 彩色图像](../../images/darkGray_32.gif)|32 及以上|
|中灰色 32|158、158、158|#9E9E9E|![中灰色 32 彩色图像](../../images/mediumGray_32.gif)|32 及以上|
|浅灰色所有|179、179、179|#B3B3B3|![浅灰色所有彩色图像](../../images/lightGray_all.gif)|所有大小|
|深灰色 16|114、114、114|#727272|![深灰色 16 彩色图像](../../images/darkGray_16.gif)|16 及以下|
|中灰色 16|144、144、144|#909090|![中灰色 16 彩色图像](../../images/mediumGray_16.gif)|16 及以下|
|蓝色 32|77、130、184|#4d82B8|![蓝色 32 彩色图像](../../images/blue_32.gif)|32 及以上|
|蓝色 16|74、125、177|#4A7DB1|![蓝色 16 彩色图像](../../images/blue_16.gif)|16 及以下|
|黄色所有|234、194、130|#EAC282|![黄色所有彩色图像](../../images/yellow_all.gif)|所有大小|
|橙色 32|231、142、70|#E78E46|![橙色 32 彩色图像](../../images/orange_32.gif)|32 及以上|
|橙色 16|227、142、70|#E3751C|![橙色 16 彩色图像](../../images/orange_16.gif)|16 及以下|
|粉色所有|230、132、151|#E68497|![粉色所有彩色图像](../../images/pink_all.gif)|所有大小|
|绿色 32|118、167、151|#76A797|![绿色 32 彩色图像](../../images/green_32.gif)|32 及以上|
|绿色 16|104、164、144|#68A490|![绿色 16 彩色图像](../../images/green_16.gif)|16 及以下|
|红色 32|216、99、68|#D86344|![红色 32 彩色图像](../../images/red_32.gif)|32 及以上|
|红色 16|214、85、50|#D65532|![红色 16 彩色图像](../../images/red_16.gif)|16 及以下|
|紫色 32|152、104、185|#9868B9|![紫色 32 彩色图像](../../images/purple_32.gif)|32 及以上|
|紫色 16|137、89、171|#8959AB|![紫色 16 彩色图像](../../images/purple_16.gif)|16 及以下|


## <a name="additional-resources"></a>其他资源

- [外接程序开发的最佳做法](../overview/add-in-development-best-practices.md)
- [Excel、Word 和 PowerPoint 的外接程序命令](../design/add-in-commands.md)
