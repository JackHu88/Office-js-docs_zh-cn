## Supertip
定义丰富的工具提示（标题和说明）。 它由“[按钮](./button.md)”和“[菜单](./menu-control.md)”控件使用。 

## 子元素
|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [标题](#标题)        | 是 |   supertip 的文本。         |
|  [说明](#说明)  | 是 |  supertip 的说明。    |

## 标题
必需。SuperTip 的文本。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 [String](./resources.md#shortstrings) 元素的 [id](./resources.md) 属性的值。

## 说明
必需。SuperTip 的描述。 **resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 [String](./resources.md#longstrings) 元素的 [id](./resources.md) 属性的值。

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```