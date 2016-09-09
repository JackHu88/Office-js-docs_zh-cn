
# Labs.Components.IDynamicComponent

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

启用与动态组件的交互。

```
interface IDynamicComponent extends Labs.Core.IComponent
```


## 属性


|名称|说明|
|:-----|:-----|
| `generatedComponentTypes: string[]`|一个包含此动态组件可能会生成的组件类型的数组。|
| `maxComponents: number`|此动态组件将生成的组件的最大数量。或者，如果没有上限，则为 **Labs.Components.Infinite**。|
