﻿# ChartFill 对象（适用于 Excel 的 JavaScript API）

表示图表元素的格式填充。

## 属性

无

## Relationships
无


## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除图表元素的填充颜色。|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|将图表元素的填充格式设置为统一颜色。|

## 方法详细信息


### clear()
清除图表元素的填充颜色。

#### 语法
```js
chartFillObject.clear();
```

#### 参数
无

#### 返回
void

#### 示例

清除名为“Chart1”的图表上值坐标轴的主要网格线的线条格式。

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;   
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### setSolidColor(color: string)
将图表元素的填充格式设置为统一颜色。

#### 语法
```js
chartFillObject.setSolidColor(color);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|color|string|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|

#### 返回
void

#### 示例

将 Chart1 的背景颜色设置为红色。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

    chart.format.fill.setSolidColor("#FF0000");

    return ctx.sync().then(function() {
            console.log("Chart1 Background Color Changed.");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```