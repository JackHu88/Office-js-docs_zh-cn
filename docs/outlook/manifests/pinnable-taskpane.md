# <a name="implement-a-pinnable-taskpane-in-outlook"></a>在 Outlook 中实现可固定的任务窗格

用于外接程序命令的[任务窗格](../add-in-commands-for-outlook.md#launching-a-task-pane)用户体验形状在打开的邮件或约会的右侧，打开一个垂直任务窗格，以便用户可以在外接程序 UI 中进行更详细的交互（填充多个字段等）。查看邮件列表时，可以在阅读窗格中看到此任务窗格，从而能够快速处理邮件。

不过，默认情况下，如果用户在阅读窗格中打开了某个邮件的外接程序任务窗格，然后选择新邮件，此任务窗格会自动关闭。如果频繁使用外接程序，用户可能更倾向于让此任务窗格一直处于打开状态，这样就无需重新激活每个邮件的外接程序了。使用可固定的任务窗格，外接程序就可以让用户如愿以偿。

> **注意**：可固定任务窗格当前仅受 Outlook 2016 for Windows（版本 7628.1000 或更高版本）的支持。

## <a name="support-taskpane-pinning"></a>支持任务窗格固定

第一步是添加固定支持，此步操作是在外接程序[清单](./manifests.md)中完成。为此，请向描述任务窗格按钮的 `Action` 元素添加 [ SupportsPinning](../../../reference/manifest/action.md#supportspinning) 元素。

由于 `SupportsPinning` 元素是在 VersionOverrides v1.1 架构中进行定义，因此必须添加 v1.0 和 v1.1 架构的 [VersionOverrides](../../../reference/manifest/versionoverrides.md) 元素。

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

有关完整示例，请参阅[命令演示示例清单](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml)中的 `msgReadOpenPaneButton` 控件。

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>根据当前选择的邮件处理 UI 更新

若要根据当前项更新任务窗格的 UI 或内部变量，必须注册事件处理程序，才能收到变化通知。

### <a name="implement-the-event-handler"></a>实现事件处理程序

事件处理程序应接受一个参数，即对象文本。该对象的 `type` 属性将设为 `Office.EventType.ItemChanged`。事件调用后，`Office.context.mailbox.item` 对象已更新，以反映当前选定项。

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

### <a name="register-the-event-handler"></a>注册事件处理程序

使用 [Office.context.mailbox.addHandlerAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#addHandlerAsync) 方法注册 `Office.EventType.ItemChanged` 事件的事件处理程序。这步操作应在任务窗格的 `Office.initialize` 函数内完成。

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="additional-resources"></a>其他资源

有关实现可固定的任务窗格的示例外接程序，请参阅 GitHub 上的[命令演示](https://github.com/jasonjoh/command-demo)。