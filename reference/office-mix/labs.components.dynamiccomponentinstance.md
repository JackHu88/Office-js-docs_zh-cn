
# <a name="labs.components.dynamiccomponentinstance"></a>Labs.Components.DynamicComponentInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示动态组件的实例。

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>属性


|属性|说明|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|组件实例定义。|

## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(component: Components.IDynamicComponentInstance)`

使用 [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md) 定义创建新的动态组件实例。


### <a name="getcomponents"></a>getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

检索此动态组件创建的所有组件。

 **参数**


|参数|说明|
|:-----|:-----|
| _callback_|已检索所有组件后触发的回调函数。|

### <a name="createcomponent"></a>createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

使用动态组件作为组件库创建新组件。

 **参数**


|参数|说明|
|:-----|:-----|
| _component_|用于创建实例的组件 ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md))。|
| _callback_|创建组件后触发的回调函数。|

### <a name="close"></a>关闭

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

指示将不存在与此组件实例相关联的任何其他提交。

 **参数**


|参数|说明|
|:-----|:-----|
| _callback_|关闭实例后，将触发的回调函数。|

### <a name="isclosed"></a>isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

返回动态组件是否已关闭。如果已关闭，则返回 **true**。

