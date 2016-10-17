
# <a name="labs.registerdeserializer"></a>Labs.registerDeserializer

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

将指定的 JSON 对象反序列化为一个对象。仅限组件作者使用。

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## <a name="parameters"></a>参数


|**名称**|**Description**|
|:-----|:-----|
|json|将要反序列化的 [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)。|

## <a name="return-value"></a>返回值

返回 [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) 实例。

