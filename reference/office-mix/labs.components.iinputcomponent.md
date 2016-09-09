
# Labs.Components.IInputComponent

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

启用与输入组件的交互。

```
interface IInputComponent extends Labs.Core.IComponent
```


## 属性


|名称|说明|
|:-----|:-----|
| `maxScore: number`|输入组件最大允许的分数。|
| `timeLimit: number`|输入问题的时间限制。|
| `hasAnswer: boolean`|如果组件有一个答案，则返回 **True**。|
| `answer: any`|组件问题的答案（如有）。|
| `secure: boolean`|如果输入组件是安全的，则返回 **True**。|
