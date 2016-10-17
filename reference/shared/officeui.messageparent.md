# <a name="ui.messageparent-method"></a>UI.messageParent 方法

将对话框中的消息传送到其父页/开始页。调用此 API 的页必须与父页位于相同的域。 

## <a name="syntax"></a>语法

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|messageObject|字符串或布尔值|接受从对话框传送到外接程序的消息。|

## <a name="returns"></a>返回
void

## <a name="examples"></a>示例
有关示例，请参阅 [DisplayDialogAsync 方法](officeui.displaydialogasync.md)主题。

