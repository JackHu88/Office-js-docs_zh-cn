

# 事件

`event` 对象作为参数传递到由无用户界面命令按钮调用的外接程序函数。该对象允许外接程序确定单击了哪个按钮，并向主机发出信号说明已完成处理。

例如，考虑外接程序清单中定义的按钮，如下所示：

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

该按钮具有设置为 `id` 的 `eventTestButton` 属性，并且将调用外接程序中定义的 `testEventObject` 函数。该函数如下所示：

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|

### 成员

####  source :Object

获取调用该方法的外接程序命令按钮的标识符。

`source` 属性返回具有以下属性的对象。

| 属性 | 说明 |
| --- | --- |
| `id` | `id` 元素的 `Control` 属性的值，用于定义外接程序清单中的外接程序命令按钮。 |

当多个按钮调用同一个函数时，可以使用此值，但是您需要基于单击的按钮采取不同的操作。

##### 类型：

*   对象

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|

##### 示例

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### 方法

####  completed()

指示外接程序已完成外接程序命令按钮触发的处理。

此方法必须在由使用 `Action` 属性设置为 `xsi:type` 的 `ExecuteFunction` 元素定义的外接程序命令调用的函数的末尾调用。调用此方法会向主机客户端发出信号，指示函数已完成并且它可以清理调用该函数所涉及的任何状态。例如，如果用户在调用此方法之前关闭 Outlook，Outlook 将警告函数仍在执行。

##### 要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](./tutorial-api-requirement-sets.md)| 1.3|
|[最低权限级别](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 受限|
|适用的 Outlook 模式| 撰写或阅读|

##### 示例

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```