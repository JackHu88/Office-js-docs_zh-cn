
# Labs.Components.ComponentAttempt

 _**适用范围：** Office 相关应用 | Office 外接程序 | Office Mix | PowerPoint_

组件上用于尝试的基类。

```
class ComponentAttempt
```


## 属性


|**名称**|**说明**|
|:-----|:-----|
| `public var _componentId: string`|指定组件的 ID。|
| `public var _id: string`|相关联的实验室 ID。|
| `public var _labs: Labs.LabsInternal`|用于与基础 [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) 进行交互的实验室 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) 对象。|
| `public var _resumed: boolean`|如果实验室已恢复给定尝试的进程，则返回 **True**。|
| `public var _state: Labs.ProblemState`|尝试的当前状态由枚举 [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md) 提供。|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|与尝试相关联的值（如有）包含在 [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md) 对象中。|

## 方法




### 构造函数

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

创建 ComponentAttempt 类的一个新实例，并提供输入参数值。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _labs_|与尝试一起使用的 [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) 实例。|
| _attemptId_|与尝试关联的 ID。|
| _值_|与尝试相关联的值数组 ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md))。|

### isResumed

 `public function isResumed(): boolean`

指示实验室是否已恢复的布尔函数。如果实验室已恢复，则为 **True**。

 **参数**

无。


### resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

指示实验室已恢复给定尝试的进程并加载现有数据作为此进程的一部分。必须恢复尝试，才可以使用该尝试。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _callback_|恢复尝试后将触发的回调函数。|

### getState

 `public function getState(): Labs.ProblemState`

检索实验室的状态。

 **参数**

无。


### processAction

 `public function processAction(action: Labs.Core.IAction): void`

执行与尝试关联的操作。

 **参数**

无。


### getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

检索与尝试关联的值。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _Key_|与值映射中的值相关联的密钥。|
