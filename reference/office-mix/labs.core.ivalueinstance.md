
# <a name="labs.core.ivalueinstance"></a>Labs.Core.IValueInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

[Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) 对象实例，其中包含数值数据（如果有的话）。

```
interface IValueInstance
```


## <a name="properties"></a>属性


|||
|:-----|:-----|
| `valueId: string`|此实例所表示的值 ID。|
| `isHint: boolean`|如果此值被认为是一个提示，则布尔值为 **true**。|
| `hasValue: boolean`|如果实例信息包含该值，则布尔值为 **true**。|
| `value?: any`|值。此参数可能已设置或可能未设置，具体取决于它是否已被隐藏。|
