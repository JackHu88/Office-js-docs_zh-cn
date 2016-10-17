
# <a name="labs.core.actions.icreatecomponentoptions"></a>Labs.Core.Actions.ICreateComponentOptions

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

创建新的组件。

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## <a name="properties"></a>属性


|||
|:-----|:-----|
| `componentId: string`|调用创建组件操作的组件。|
| `component: Core.IComponent`|要创建的 [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md) 组件|
| `correlationId?: string`|在实验室的所有实例中关联此组件的可选字段。允许主机确定对同一组件的不同尝试。|
