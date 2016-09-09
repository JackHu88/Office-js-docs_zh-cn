
# Labs.Components.ActivityComponentInstance

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

表示活动组件的当前实例。

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## 属性


|**名称**|**说明**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|此类表示的基础 [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)。|

## 方法




### 构造函数

 `function constructor(component: Components.IActivityComponentInstance)`

创建 [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) 类的新实例。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _组件_|将从该类创建此类的 **IActivityComponentInstance**。|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

生成一个新的 **ActivityComponentAttempt** 实例，并实现在基类上定义的抽象方法

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _createAttemptResult_|创建尝试操作的结果。|
