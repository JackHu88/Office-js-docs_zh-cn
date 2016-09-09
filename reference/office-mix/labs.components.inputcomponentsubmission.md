
# Labs.Components.InputComponentSubmission

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

表示提交到输入组件。

```
class InputComponentSubmission
```


## 属性


|属性|说明|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|与提交相关联的答案 ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md))。|
| `public var result: Components.InputComponentResult`|提交的结果 ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md))。|
| `public var time: number`|收到提交的时间。|

## 方法




### 构造函数

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

创建 **InputComponentSubmission** 类的新实例。

 **参数**


|参数|说明|
|:-----|:-----|
| _answer_|与提交相关联的答案。|
| _result_|提交的结果。|
| _时间_|收到提交的时间。|
