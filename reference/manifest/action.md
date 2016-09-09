# Action 元素
 指定用户选择 [按钮](./button-control.md) 或 [菜单](./menu-control.md) 控件时将执行的操作。
 
## 属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要执行的操作类型|


## 子元素

|  元素 |  说明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要执行的函数的名称。 |
|  [SourceLocation](#sourcelocation) |    指定该操作的源文件位置。 |
  

## xsi:type
此属性指定当用户选择按钮时所执行的操作类型。 它可以是下列值之一：
- ExecuteFunction
- ShowTaskpane

## FunctionName
**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](./functionfile.md) 元素指定的文件中。

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## SourceLocation
**xsi:type** 为 ShowTaskpane 时的必需元素。指定此操作的源文件位置。 **resid** 属性必须设置为 **Urls** 元素（位于 **Resources** 元素）中 [Url](./resources.md#urls) 元素的 [id](./resources.md) 属性的值。

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  
