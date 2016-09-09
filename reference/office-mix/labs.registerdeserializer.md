
# Labs.registerDeserializer

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

将指定 JSON 对象反序列化为一个对象。仅供组件作者使用。

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## 参数


|**Name**|**说明**|
|:-----|:-----|
|json|将要反序列化的 [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)。|

## 返回值

返回 [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) 实例。

