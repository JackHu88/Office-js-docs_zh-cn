#<a name="ui.dialog-object"></a>UI.Dialog object
调用 [displayDialogAsync](officeui.displaydialogasync.md) 方法时返回的对象。

## <a name="members"></a>成员
| 成员       | 类型   |说明|
|:---------------|:--------|:----------|
|关闭|函数|允许外接程序关闭其对话框。|
|addEventHandler|函数|注册事件处理程序。支持两种类型的事件： <ul><li>DialogMessageReceived。在对话框向其父级发送消息时触发。</li><li>DialogEventReceived。在对话框已关闭或以其他方式卸载时触发。</li></ul> |


### <a name="close()"></a>close()
从父页调用以关闭相应的对话框。     
```js    
[dialogObject].close();    
``` 

#### <a name="parameters"></a>参数    
无。 

#### <a name="returns"></a>返回    
void  


#### <a name="examples"></a>示例
有关示例，请参阅 [DisplayDialogAsync 方法](officeui.displaydialogasync.md)主题。
