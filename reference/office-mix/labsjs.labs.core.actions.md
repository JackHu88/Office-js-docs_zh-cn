
# LabsJS.Labs.Core.Actions
提供了 LabJS.Labs.Core.Actions JavaScript API 的概述。

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

这些 API 表示指明实验室当前行为的实验室操作。这些 API 对创建新组件或使用新驱动程序（而非 Office Mix）建立连接非常有用。

## LabsJS.Labs.Core.Actions API 模块

Actions 模块包含以下类型：


### 接口


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|要关闭的组件。|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|与尝试关联的组件。|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|创建给定组件尝试的结果。|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|创建新的组件。|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|创建新组件的 [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) 结果。|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|获取值操作的结果。|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|提交尝试答案的结果。|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|当前尝试的超时操作可用的选项。|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|获取值操作时可用的选项。|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|与恢复尝试关联的选项。|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|提交答案操作时可用的选项。|

### 变量


|||
|:-----|:-----|
| `var CloseComponentAction: string`|关闭组件并指示不再会对其进行更多操作。|
| `var CreateAttemptAction: string`|创建新尝试的操作。|
| `var CreateComponentAction: string`|创建新组件的操作。|
| `var AttemptTimeoutAction: string`|尝试超时操作。|
| `var GetValueAction: string`|检索与尝试关联的值的操作。|
| `var ResumeAttemptAction: string`|尝试超时操作。|
| `var SubmitAnswerAction: string`|针对给定的尝试提交答案的操作。|
