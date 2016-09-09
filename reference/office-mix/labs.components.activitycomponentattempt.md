
# Labs.Components.ActivityComponentAttempt

 _**适用范围：** Office 相关应用 | Office 外接程序 | Office Mix | PowerPoint_

表示尝试完成活动组件。

```
class Permissions
```


## 方法




### 构造函数

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

创建 **ActivityComponentAttempt** 类的新实例。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _labs_|与组件相关联的实验室实例 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentId_|与 attempt 关联的组件 ID。|
| _attemptId_|尝试的 ID。|
| _值_|与组件关联的值（如有）。|

### complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

指示活动已完成的指示器。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _callback_|活动完成后调用的回调函数。|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

运行检索有关给定尝试的操作的函数，然后填写实验室的状态。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _action_|操作实例 ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md))。|
