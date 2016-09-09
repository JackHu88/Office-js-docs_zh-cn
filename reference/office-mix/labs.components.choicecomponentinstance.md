
# Labs.Components.ChoiceComponentInstance

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

表示选项组件的实例。

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## 属性


|属性|说明|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|此类表示的基础 [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md)。|

## 方法




### 构造函数

 `function constructor(component: Components.IChoiceComponentInstance)`

创建 **ChoiceComponentInstance** 类的新实例。

 **参数**


|参数|说明|
|:-----|:-----|
| _组件_|用来创建此类的 [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) 对象。|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

生成一个新的 **ChoiceComponentAttempt** 实例，并实现在基类上定义的抽象方法。

 **参数**


|参数|说明|
|:-----|:-----|
| _createAttemptResult_|创建尝试操作产生的结果。|
