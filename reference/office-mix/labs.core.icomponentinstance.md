
# Labs.Core.IComponentInstance

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

实验室组件实例的基类。

```
interface IComponentInstance extends Core.ILabObject, Core.IUserData
```


## 属性


|||
|:-----|:-----|
| `componentId: string`|与此实例关联的组件 ID。|
| `name: string`|组件的名称。|
| `values: {[type:string]: Core.IValueInstance[]}`|与组件相关联的值属性映射。|

## 注解

组件实例是对用户组件的实例化。它包含有关实验室的特定运行组件的翻译视图。此视图可能排除了隐藏的信息（答案、提示等），还包含用来标识各种实例的 ID。

