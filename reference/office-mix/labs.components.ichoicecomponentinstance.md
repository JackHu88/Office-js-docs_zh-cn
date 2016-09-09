
# Labs.Components.IChoiceComponentInstance

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

选项组件的实例。

```
interface IChoiceComponentInstance extends Labs.Core.IComponentInstance
```


## 属性


|名称|说明|
|:-----|:-----|
| `choices: Components.IChoice[]`|一个表示与此问题相关联的选项列表的数组。|
| `timeLimit: number`|完成该问题的时间限制。|
| `maxAttempts: number`|允许针对该问题尝试的最大数目。|
| `maxScore: number`|该问题的最大分数。|
| `hasAnswer: boolean`|如果该问题有答案，则返回 **True**。|
| `answer: any`|问题的答案。如果支持多种答案，即为数组，或者如果只支持一种答案，则为单个 ID。|
| `secure: boolean`|无论测验是否安全，都意味着从用户中保留安全字段。|
