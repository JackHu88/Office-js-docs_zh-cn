# <a name="outlook-add-in-design-guidelines"></a>Outlook 外接程序设计准则

外接程序是合作伙伴扩展 Outlook 核心功能集之外的功能的好方法。用户可以通过外接程序访问第三方体验、任务和内容，而无需离开其收件箱。安装后，Outlook 外接程序可用于所有平台和设备。以下高级指南将有助于设计和生成引人注目的外接程序，其可将应用的最佳功能直接引入 Windows、Web、iOS、Mac 和 Android 上的 Outlook（即将推出）。

## <a name="principles"></a>原则

1. **重点关注几个关键任务；并将其做好**

    设计最佳的外接程序易于使用、目标明确并且可为用户带来实际价值。由于外接程序将在 Outlook 内部运行，因此这一原则额外重要。Outlook 是生产力应用 – 人们使用此应用来完成工作。

    你将成为我们体验的扩展测试人员，请务必确保启用方案就像是在 Outlook 内部进行操作一样自然恰当。认真考虑你的哪些常用用例通过与这些方案挂钩可以从我们的电子邮件和日历体验中获益最大。

    外接程序不应尝试执行应用所执行的一切操作。重点应放在 Outlook 内容的上下文中使用最频繁的恰当操作。考虑操作调用并明确任务窗格打开时用户应执行什么操作。

2. **使其尽可能类似于本机模式**

    应使用正在运行 Outlook 的平台上的本机模式设计外接程序。若要实现这一点，务必尊重并实现各个平台规定的交互和外观准则。Outlook 具有自己的准则，同样也必须考虑这些准则。设计良好的外接程序将恰当地融合体验、平台和 Outlook。

    这就是说，外接程序在 Outlook for iOS 和 Outlook for Android（如果我们提供支持的话）上运行时的外观必须不同。我们建议不妨使用 [Framework7](https://framework7.io/) 作为样式设置选项。随着我们逐步提供适用于 Outlook for Android 的外接程序支持，我们将发布更新后的指南，特别是针对 Android。

3. **确保使用体验令人满意，并正确设置详细信息**

    人们喜欢使用实用且外观吸引人的产品。在已仔细考虑每个交互和外观细节的情况下精心创建体验有助于确保外接程序成功。完成任务的必要步骤必须清楚并相互关联。理想情况下，操作不应超过一次或两次单击。尝试不要使用户脱离上下文来完成操作。用户应可以轻松进入和退出外接程序并可轻松返回至用户之前正在执行的操作。外接程序不应成为花费大量时间的目标，它是对核心功能的增强。如果处理得当，外接程序将有助于实现使用户更高效的目标。

4. **明智地进行品牌打造**

    我们非常重视品牌打造，同时我们知道向用户提供唯一体验至关重要。但是我们认为确保外接程序成功的最佳方式是生成巧妙整合品牌元素的直观体验，而非显示重复或突兀的品牌元素，它们只会分散用户无阻碍进入系统的注意力。有效地整合品牌的良好方式是使用品牌颜色、图标和声音（假定这些与首选的平台模式或辅助功能要求不冲突）。努力将重点集中于内容和任务完成，而非品牌关注。

## <a name="design-patterns"></a>设计模式

> **注意：**上述准则适用于所有终结点/平台，但以下模式和示例特定于 iOS 平台上的移动外接程序。

我们提供了包含适用于 Outlook Mobile 环境的 iOS 移动模式的[模板](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/tree/master/Helpful%20Templates/Outlook%20Mobile)，以帮助创建设计良好的外接程序。利用这些特定模式有助于确保外接程序如同在 iOS 平台和 Outlook Mobile 本机自带一般。下面详细介绍了这些模式。虽不全面，但这只是构建库的开始，在我们发现合作伙伴希望纳入其外接程序的其他范例时我们将继续构建此库。  

### <a name="overview"></a>概述

典型的外接程序由下列组件组成。

![iOS 上的任务窗格的基本 UX 模式关系图](../../images/outlook-mobile-design-overview.png)

### <a name="loading"></a>加载

用户点击外接程序后，UX 应尽快显示。如果出现任何延迟，则使用进度栏或活动指示器。时间量可确定时应使用进度栏，时间量不可确定时应使用活动指示器。

![iOS 上的进度栏和活动指示器示例](../../images/outlook-mobile-design-loading.png)

### <a name="sign-insign-up"></a>登录/注册

使登录（和注册）流简单且易用使用。

![iOS 上的登录和注册页示例](../../images/outlook-mobile-design-signin.png)

### <a name="brand-bar"></a>品牌栏

外接程序的第一个屏幕应包含品牌元素。品牌栏用于进行识别，同时也有助于为用户设置上下文。由于导航栏包含公司/品牌的名称，因此没有必要在后续页面上重复品牌栏。

![iOS 上的品牌栏示例](../../images/outlook-mobile-design-branding.png)

### <a name="margins"></a>边距

移动电话边距每侧应设置为 15px（屏幕的 8%），与 Outlook iOS 一致。

![iOS 上的边距示例](../../images/outlook-mobile-design-margins.png)

### <a name="typography"></a>版式

版式使用与 Outlook iOS 对齐并尽量简单以保证可扫描性。

![适用于 iOS 的版式示例](../../images/outlook-mobile-design-typography.png)

### <a name="color-palette"></a>调色板

颜色使用在 Outlook iOS 中比较微妙。我们要求颜色使用本地化到操作和错误状态，以保证一致，仅品牌栏使用唯一的颜色。

![适用于 iOS 的调色板](../../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a>单元格

由于导航栏不能用于标记页面，因此使用节标题标记页面。

![适用于 iOS 的单元格类型](../../images/outlook-mobile-design-cell-types.png)
* * *
![适用于 iOS 的单元格“待办事项”](../../images/outlook-mobile-design-cell-dos.png)
* * *
![适用于 iOS 的单元格“禁止事项”](../../images/outlook-mobile-design-cell-donts.png)
* * *
![适用于 iOS 的单元格和输入](../../images/outlook-mobile-design-cell-input.png)

### <a name="actions"></a>操作

即使应用要处理大量操作，也要考虑想要外接程序执行的最重要的操作，并重点关注这些操作。

![iOS 中的操作和单元格](../../images/outlook-mobile-design-action-cells.png)
* * *
![适用于 iOS 的操作“待办事项”](../../images/outlook-mobile-design-action-dos.png)

### <a name="buttons"></a>按钮

存在以下其他 UX 元素时使用按钮（vs. 操作，其中操作是屏幕上的最后一个元素）。

![适用于 iOS 的按钮示例](../../images/outlook-mobile-design-buttons.png)

### <a name="tabs"></a>选项卡

选项卡有助于内容组织。

![适用于 iOS 的选项卡示例](../../images/outlook-mobile-design-tabs.png)

### <a name="icons"></a>图标

图标应尽可能遵循当前 Outlook iOS 的设计。使用标准大小和颜色。

![适用于 iOS 的图标示例](../../images/outlook-mobile-design-icons.png)

## <a name="end-to-end-examples"></a>端到端示例

为了推动 v1 Outlook Mobile 外接程序的启动，我们已与正在生成外接程序的合作伙伴密切合作。作为展示其外接程序在 Outlook Mobile 上的潜力的方式，我们的设计人员使用我们的准则和模式将每个外接程序的端到端流组合在一起。

> **重要说明：**这些示例旨在强调同时处理外接程序的交互和外观设计的理想方法，可能与外接程序发布版本中的准确功能集不匹配。 

### <a name="giphy"></a>GIPHY

![适用于 GIPHY 外接程序的端到端设计](../../images/outlook-mobile-design-giphy.png)

### <a name="nimble"></a>Nimble

![适用于 Nimble 外接程序的端到端设计](../../images/outlook-mobile-design-nimble.png)

### <a name="trello"></a>Trello

![适用于 Trello 外接程序的端到端设计第 1 部分](../../images/outlook-mobile-design-trello-1.png)
* * *
![适用于 Trello 外接程序的端到端设计第 2 部分](../../images/outlook-mobile-design-trello-2.png)
* * *
![适用于 Trello 外接程序的端到端设计第 3 部分](../../images/outlook-mobile-design-trello-3.png)

### <a name="dynamics-crm"></a>Dynamics CRM

![适用于 Dynamics CRM 外接程序的端到端设计](../../images/outlook-mobile-design-crm.png)
