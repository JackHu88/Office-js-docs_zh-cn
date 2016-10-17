
# <a name="labs.components.idynamiccomponentinstance"></a>Labs.Components.IDynamicComponentInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

动态组件的实例。

```
interface IDynamicComponentInstance extends Labs.Core.IComponentInstance
```


## <a name="properties"></a>属性


|名称|说明|
|:-----|:-----|
| `generatedComponentTypes: string[]`|一个包含此动态组件可能会生成的组件类型的数组。|
| `maxComponents: number`|此动态组件将生成的组件的最大数量。或者，如果没有上限，则为 **Labs.Components.Infinite**。|
