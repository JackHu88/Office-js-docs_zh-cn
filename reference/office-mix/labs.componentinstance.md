
# <a name="labs.componentinstance"></a>Labs.ComponentInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示组件的一个实例，这是运行时对用户的指定组件的实例化。对象包含实验室特定运行的组件的已翻译视图。

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>属性

无。


## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor()`

初始化 **ComponentInstance** 类的新实例。


### <a name="createattempt"></a>createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

在组件的上下文中创建新尝试。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _callback_|创建尝试后触发的回调。|

### <a name="getattempts"></a>getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

检索与指定组件关联的所有尝试。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _callback_|检索尝试后触发的回调。|

### <a name="getcreateattemptoptions"></a>getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

说明


### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

从给定的操作生成尝试。应由派生类实现。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _createAttemptResult_|指定尝试的创建尝试操作。|
