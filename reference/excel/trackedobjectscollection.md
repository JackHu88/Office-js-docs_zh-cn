# TrackedObjectsCollection 对象（适用于 Office 2016 的 JavaScript API）

允许外接程序管理各 sync() 批次的 range 对象引用。通常情况下，Excel.run() 允许您自动维护各批次的引用，而无需显式跟踪它们。但是，如果某个外接程序方案要求跟踪并手动调整 range 对象以反映基础 Excel 区域的当前状态，该集合可用于标记此类要跟踪的对象。请注意，如果某个 range 对象标记为要进行跟踪，则需在使用时显式删除以释放 Excel 中的内存，尤其是在出现错误时。

## 属性
无。

## Relationships

无

## 方法

trackedObjectsCollection 对象具有下列定义的方法：

| 方法     | 返回类型    |说明|
|:-----------------|:--------|:----------|
|[add(rangeObject:Range)](#addrangeobject-range)| Null             |创建对区域的新引用。|
|[remove(rangeObject:Range)](#removerangeobject-range)| Null             |删除对区域的引用。  |
|[removeAll()](#removeallrangeobject-range)| Null|删除外接程序在设备上创建的所有引用。|


## API 规范 

### add(rangeObject: range)
向 trackedObjectsCollection 添加一个 range 对象。将会跟踪跨批次请求的任何基础变更，任何后续更新将应用到 range 对象的当前状态。 

#### 语法
```js
trackedObjectsCollection.add(rangeObject);
```

#### 参数

参数       | 类型   | 说明
--------------- | ------ | ------------
`rangeObject`  | [Range 对象设置内联图片](range.md)| 需添加到 trackedObjectCollection 的 Range 对象。

#### 返回
Null

#### 示例

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    return ctx.sync(); 
});
```


### remove(rangeObject: range)

从集合中删除引用对象。这可以释放维护所跟踪对象的状态所需的内存和资源。请注意，如果某个 range 对象标记为要跟踪，则即使在出现错误时也需显式删除。

#### 语法
```js
trackedObjectsCollection.remove(rangeObject);
```

#### 参数

参数       | 类型   | 说明
--------------- | ------ | ------------
`rangeObject`  | [Range 对象设置内联图片](range.md)| 需从 trackedObjectCollection 中删除的 range 对象。

#### 返回
Null

#### 示例


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.remove(range); 
    return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

删除外接程序在设备上创建的所有引用。

#### 语法
```js
trackedObjectsCollection.removeAll();
```

#### 参数

无

#### 返回
Null

#### 示例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var ctx = new Excel.RequestContext();
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    ctx.trackedObjectsCollection.add(range);
    ctx.load(range);
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.removeAll(); 
    return ctx.sync(); 
});
```
