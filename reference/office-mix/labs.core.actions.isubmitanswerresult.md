
# Labs.Core.Actions.ISubmitAnswerResult

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

提交尝试答案的结果。

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## 属性


|||
|:-----|:-----|
| `submissionId: string`|与提交相关的 ID。由服务器提供。|
| `complete: boolean`|如果由于当前提交而完成尝试，则返回 **true**。|
| `score: any`|与提交关联的分数信息。|
