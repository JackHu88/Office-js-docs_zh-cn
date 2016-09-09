
# Labs.Core.IComponent

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

表示实验室组件的基类。

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## 属性


|||
|:-----|:-----|
| `name: string`|组件的名称。|
| `values: {[type:string]: Core.IValue[]}`|与组件相关联的值属性映射。|
