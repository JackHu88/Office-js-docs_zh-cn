
# <a name="labs.components.inputcomponentresult"></a>Labs.Components.InputComponentResult

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

输入组件提交的结果。

```
class InputComponentResult
```


## <a name="properties"></a>属性


|属性|说明|
|:-----|:-----|
| `public var score: any`|与提交相关联的分数。|
| `public var complete: boolean`|指示提交的结果是否会导致完成尝试。如果尝试已完成，则返回 **True**。|

## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(score: any, complete: boolean)`

创建 **InputComponentResult** 类的新实例。

 **参数**


|参数|说明|
|:-----|:-----|
| _score_|与结果关联的分数。|
| _complete_|如果结果完成了尝试，则布尔值为 **true**。|
