
# Labs.Components.ChoiceComponentAttempt

 _**适用范围：** Office 相关应用 | Office 外接程序 | Office Mix | PowerPoint_

表示选项组件上的尝试。

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## 方法




### 构造函数

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

创建 **ChoiceComponentAttempt** 类的新实例。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _labs_|与尝试一起使用的 [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) 实例。|
| _attemptId_|与尝试关联的 ID。|
| _值_|与尝试关联的值。|

### timeout

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

指示实验室已超时。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _callback_|服务器已收到超时消息后，将触发的回调函数。|

### getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

检索之前针对给定尝试提交的所有提交。


### submit

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

提交一个由实验室评分的新答案，并且不使用主机来计算成绩。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _answer_|尝试的答案。|
| _result_|提交的结果。|
| _callback_|接收到提交后触发的回调函数。|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

开始处理 [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md) 操作。

