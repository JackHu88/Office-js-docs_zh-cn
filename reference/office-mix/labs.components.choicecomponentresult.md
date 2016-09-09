
# Labs.Components.ChoiceComponentResult

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

选项组件提交的结果。

```
class ChoiceComponentResult
```


## 属性


|属性|说明|
|:-----|:-----|
| `public var score: any`|与提交相关联的分数。|
| `public var complete: boolean`|结果是否已完成了尝试。如果结果完成了尝试，则返回 **True**。|

## 方法




### 构造函数

 `function constructor(score: any, complete: boolean)`

创建 **ChoiceComponentResult** 类的新实例。

 **参数**


|参数|说明|
|:-----|:-----|
| _分数_|结果的分数。|
| _complete_|指示结果是否完成了尝试。|
