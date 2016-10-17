
# <a name="labs.components.inputcomponentinstance"></a>Labs.Components.InputComponentInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示输入组件的实例。

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## <a name="properties"></a>属性


|属性|说明|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|由此类表示的基础 [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) 对象。|

## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(component: Components.IInputComponentInstance)`

创建一个新的 [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) 实例。

 **参数**


|参数|说明|
|:-----|:-----|
| _component_|用来创建此类的 [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)。|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

生成一个新的 [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md)。实现在基类上定义的抽象方法。

 **参数**


|参数|说明|
|:-----|:-----|
| _createAttemptResult_|创建尝试操作的结果。|
