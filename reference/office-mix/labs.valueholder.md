
# <a name="labs.valueholder"></a>Labs.ValueHolder

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

保留和跟踪指定实验室的值的容器对象。值可能存储在本地或服务器上。

```
class ValueHolder<T>
```


## <a name="variables"></a>变量


|||
|:-----|:-----|
| `public var isHint: boolean`|如果值为提示，则返回 **True**。|
| `public var hasBeenRequested: boolean`|如果值被实验室请求，则返回 **True**。|
| `public var hasValue: boolean`|如果值容器当前具有所需的值，则返回 **True**。|
| `public var value: T`|保存在容器中的值。|
| `public var id: string`|ID 值。|

## <a name="methods"></a>方法




### <a name="getvalue"></a>getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

检索指定的值。

 **参数**


|||
|:-----|:-----|
| _callback_|返回指定值的回调函数。|

### <a name="providevalue"></a>provideValue

 `public function provideValue(value: T): void`

提供值到值容器的内部方法。

 **参数**


|||
|:-----|:-----|
| _value_|要向值容器提供的值。|
