
# <a name="labs.core.actions.isubmitansweroptions"></a>Labs.Core.Actions.ISubmitAnswerOptions

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

提交答案操作时可用的选项。

```
interface ISubmitAnswerOptions extends Core.IActionOptions
```


## <a name="properties"></a>属性


|||
|:-----|:-----|
| `componentId: string`|与提交关联的组件。|
| `attemptId: string`|与提交关联的尝试。|
| `answer: any`|正在提交的答案。|
