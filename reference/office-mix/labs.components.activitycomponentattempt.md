
# <a name="labs.components.activitycomponentattempt"></a>Labs.Components.ActivityComponentAttempt

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示尝试完成活动组件。

```
class Permissions
```


## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

创建 **ActivityComponentAttempt** 类的新实例。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _labs_|与组件相关联的实验室实例 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentId_|与尝试关联的组件 ID。|
| _attemptId_|尝试的 ID。|
| _values_|与组件关联的值（如有）。|

### <a name="complete"></a>complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

指示活动已完成的指示器。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _callback_|活动完成后调用的回调函数。|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

运行检索有关给定尝试的操作的函数，然后填写实验室的状态。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _action_|操作实例 ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md))。|
