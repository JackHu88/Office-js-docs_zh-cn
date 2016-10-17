
# <a name="labs.components.inputcomponentattempt"></a>Labs.Components.InputComponentAttempt

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示尝试与输入组件交互。

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

创建 **InputComponentAttempt** 类的新实例。

 **参数**


|参数|说明|
|:-----|:-----|
| _labs_|与尝试相关联的实验室 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentID_|与尝试关联的组件 ID。|
| _attemptId_|特定尝试的 ID。|
| _values_|包含值实例 ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)) 的数组。|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

循环访问指定尝试的检索操作并填充实验室的状态。

 **参数**


|参数|说明|
|:-----|:-----|
| _action_|与实验室状态相关联的操作。|

### <a name="getsubmissions"></a>getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

检索所有先前已提交的针对指定尝试的提交。


### <a name="submit"></a>submit

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

提交一个由实验室评分的新答案，并且不使用主机来计算成绩。

 **参数**


|参数|说明|
|:-----|:-----|
| _answer_|与尝试相关联的答案。|
| _result_|与提交关联的结果。|
| _callback_|接收到提交后触发的回调函数。|
