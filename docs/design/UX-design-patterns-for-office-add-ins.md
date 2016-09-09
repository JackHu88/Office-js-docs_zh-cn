# 适用于 Office 外接程序的 UX 设计模式。 

在设计时 Office 外接程序时，外接程序的 UX 设计应该提供极具吸引力的扩展 Office 体验。若要创建一个极好的外接程序，外接程序应该提供首次运行体验、一流的 UX 体验以及页面之间与其他功能之间的平稳过渡。提供整洁、现代的 UX 体验增加用户保留时间和外接程序的采用率。本文为以下设计人员和开发人员提供了 UX 资源：

* 介绍了基于最佳实践的常见 UX 设计模式。
* 实现 Office 结构组件和样式。
* 实现类似于默认 Office UI 自然延伸的外接程序。 

## 如何开始使用 Office 外接程序设计示例资源？

使用这些设计或代码资产没有任何前提条件。开始创建外接程序的良好 UX：

* 查看 UX 设计模式，并确定哪些模式对您的外接程序非常重要。例如，选取首次运行体验之一。
* 然后执行下列一项或多项操作：
	* 将代码文件复制到外接程序项目，并开始自定义这些代码文件以满足您的需求。您将需要 [common.js 文件](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/)、[assets 文件夹](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets)，以及适用于您需要的设计模式的代码文件夹。请参阅下面的链接。
	* 下载参考 PDF 并将其用作创建自己的 UX 设计时的指南。请参阅下面的链接。
	* 下载 Adobe Illustrator 文件并编辑这些文件以模仿您自己的外接程序设计。从[此处](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files)获取文件。
 

## 首次运行

首次运行体验是用户第一次打开您的外接程序时获得的体验。下面列出了可包含在外接程序中的首次运行设计模式。下面列出了每个设计模式的图像。

* **开始步骤**为用户提供执行步骤的排序列表以开始使用您的外接程序。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/instruction-step)）
* **值**传达您的外接程序的价值主张。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/value-placemat)）
* **视频**在用户开始使用您的外接程序之前向其展示视频。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/video-placemat)）
* **演练**让用户在开始使用外接程序之前熟悉一系列功能或信息。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/walkthrough)）
* [Office 应用商店](https://msdn.microsoft.com/zh-cn/library/office/jj220033.aspx)有一个系统来为用户提供外接程序的试用版，但如果希望完全控制 UI 的试用体验，请使用以下模板：
	* **试用版**演示用户如何开始使用外接程序的试用版。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat)）
	* **试用版功能**提醒用户他们尝试使用的功能不可在外接程序的试用版中使用。（[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat-feature)）


> 注意:考虑一次或多次向用户显示首次运行体验是否对您的方案非常重要。例如，如果用户定期使用您的外接程序，他们可能会忘记如何使用外接程序。再次看到首次运行体验可能对一些用户非常有帮助。 

 <table>
 <tr><th>开始步骤</th><th>值</th><th>视频</th></tr>
 <tr><td>![instruction steps" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/instruction.step.PNG)</td><td>![value placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/value.placemat.PNG)</td><td>![video placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/video.placemat.PNG)</td></tr>
 </table>

 <table>
 <tr><th>演练第一页</th><th>试用版</th><th>试用版功能</th></tr>
 <tr><td>![walkthrough 1" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/walkthrough1.PNG)</td><td>![trial placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.PNG)</td><td>![trial placemat feature" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.feature.PNG)</td></tr>
 </table> 


## 泛型和品牌打造

* **登录页**是用户在首次运行体验或登录流程后导航到的首个位置。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/generic/landing-page)）

<table>
 <tr><th>登录</th></tr>
 <tr><td>![landing page" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/landing.page.PNG)</td></tr>
 </table>

## 通知

外接程序可以通过多个方法向用户通知事件，如错误、或进度。下面列出了这些技术。下面列出了每个技术的图像。

* **嵌入式对话框**显示在任务窗格中使用按钮或其他控件提供信息或互动体验的对话框。请考虑使用其中之一提示用户确认操作。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/embedded-dialog)）
* **内联消息**表示错误、成功或信息，它可以出现在任务窗格中的指定位置。例如，如果用户在文本框中输入格式不正确的电子邮件地址，文本框下方将出现一条错误消息。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/inline-message)）
* **消息横幅**在可折叠为一行、扩展到多行或解除的横幅中提供信息或简单调用操作。考虑使用消息横幅来在外接程序启动时报告服务更新或有用的提示。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_messagebanner.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/message-banner)）
* **进度栏**表示长期运行的同步过程（例如，用户在执行任何进一步操作前必须完成的配置任务）的进度。这是一个加强外接程序品牌的单独间隙页面。在过程可发送返回到外接程序的进度的定期度量值时，使用进度栏。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/progress-bar)）
* **微调框**表示一个长时间运行的同步过程正在进行，但不提供这一过程的进度。这是一个加强外接程序品牌的单独间隙页面。在外接程序无法知晓某一过程的可靠进度时，使用微调框。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/spinner)）
* **Toast** 提供一个会在几秒钟后消失的简短信息。由于用户可能看不到该消息，toast 仅用作非基本信息。在远程系统中这是通知用户某个事件的理想选择，如收到一封电子邮件。（[PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_toast.pdf "PDF")、[代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/toast)）

 <table>
 <tr><th>嵌入式对话框</th><th>内联消息</th><th>消息横幅</th></tr>
 <tr><td>![embedded dialog" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/embedded.dialog.PNG)</td><td>![inline message" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/inline.message.PNG)</td><td>![message banner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/message.banner.PNG)</td></tr>
 </table>

 <table>
 <tr><th>进度栏</th><th>微调框</th><th>Toast</th></tr>
 <tr><td>![progress bar" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/progress.bar.PNG)</td><td>![spinner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/spinner.PNG)</td><td>![toast" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/toast.PNG)</td></tr>
 </table>

## 已知问题

* 运行外接程序项目外的某些代码文件会引发 JavaScript 错误。 
	* 解决方案：确保将这些文件添加到 Office 外接程序项目。 
	
## 其他资源

* [开发 Office 外接程序的最佳做法](https://dev.office.com/docs/add-ins/design/add-in-development-best-practices)
* [Office UI 结构](http://dev.office.com/fabric/)
