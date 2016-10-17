
# <a name="labs.labinstance"></a>Labs.LabInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

为当前用户配置的实验室的实例。使用此对象可记录和检索用户的实验室数据。

```
class LabInstance
```


## <a name="variables"></a>变量


|||
|:-----|:-----|
| `public var data: any`|存放用户数据的容器变量。|
| `public var components: Labs.ComponentInstanceBase[]`|构成实验室实例的组件。|

## <a name="methods"></a>方法




### <a name="getstate"></a>getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

针对给定用户检索实验室的当前状态。

 **参数**


|||
|:-----|:-----|
| _callback_|检索实验室状态时触发的回调函数。|

### <a name="setstate"></a>setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

针对给定用户设置实验室的状态。

 **参数**


|||
|:-----|:-----|
| _state_|要设置的状态。|
| _callback_|设置状态后触发的回调函数。|

### <a name="done"></a>Done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

指示用户已完成使用实验室操作的指示函数。

 **参数**


|||
|:-----|:-----|
| _callback_|完成实验室操作后触发的回调函数。|
