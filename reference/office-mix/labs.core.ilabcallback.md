
# Labs.Core.ILabCallback

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

用于处理 Labs.js 回调方法的接口。

```
interface ILabCallback<T>
```


## 回调签名

 `(err: any, data: T): void`

 **回调参数**


|||
|:-----|:-----|
| _err_|如果没有错误发生，则返回 **Null**。如果发生了错误，则返回非 **null**。|
| _data_|使用回调返回的数据。|
