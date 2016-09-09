# Group 元素
定义选项卡中的一组 UI 扩展点。  在自定义选项卡上，外接程序可以创建最多 10 个组。 每个组限制为 6 个控件，不论它显示在哪个选项卡上。 外接程序限定到一个自定义选项卡。

## 属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [id](#id)  |  是  | 组的唯一 ID。|

## 子元素
|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [TermStore](#termstore)      | 是 |  CustomTab 或组的标签。  |
|  [控件](#控件)    | 是 |  一个或多个控件对象的集合。  |

## id attribute
必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。

## 标签 
必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 [String](./resources.md#shortstrings) 元素的 [id](./resources.md) 属性的值。

## 控件
一个组需要至少一个控件。 目前，仅支持“[按钮](./control.md#button-control)”和“[菜单](./menu.md#menu-control)”。 

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```