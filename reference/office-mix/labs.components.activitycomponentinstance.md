
# <a name="labs.components.activitycomponentinstance"></a>Labs.Components.ActivityComponentInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示活动组件的当前实例。

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## <a name="properties"></a>属性


|**名称**|**Description**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|此类表示的基础 [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)。|

## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(component: Components.IActivityComponentInstance)`

创建 [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) 类的新实例。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _component_|将从该类创建此类的 **IActivityComponentInstance**。|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

生成一个新的 **ActivityComponentAttempt** 实例，并实现在基类上定义的抽象方法

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _createAttemptResult_|创建尝试操作的结果。|
