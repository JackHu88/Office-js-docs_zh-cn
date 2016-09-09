
# Labs.ComponentInstance

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

表示组件的一个实例，这是运行时对用户的指定组件的实例化。对象包含实验室特定运行的组件的已翻译视图。

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## 属性

无。


## 方法




### 构造函数

 `function constructor()`

初始化 **ComponentInstance** 类的新实例。


### createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

在组件的上下文中创建新尝试。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _callback_|创建尝试后触发的回调。|

### getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

检索与指定组件关联的所有尝试。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _callback_|检索尝试后触发的回调。|

### getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

说明


### buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

检索创建尝试的默认选项。可以被派生类重写。

 **参数**


|**Name**|**说明**|
|:-----|:-----|
| _createAttemptResult_|指定尝试的创建尝试操作。|
