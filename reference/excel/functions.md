# 工作表函数
用作可从 JavaScript 或 REST 中调用的 Microsoft Excel 工作表函数的容器。

## 返回类型
工作表函数返回 FunctionResult 对象。FunctionResult 对象具有两个属性。

| 属性       | 类型    |说明|注释 |
|:---------------|:--------|:----------|:-----|
|value|string|返回应用指定的工作表函数的结果。||
|错误|string|返回应用指定的工作表函数时的错误信息。||


## 语法
```js
workbook.functions.functionMethod();
```

## 示例
下面是跟踪 Excel 中不同工具销售情况的示例表。我们将使用此表中的数据解释工作表函数的工作方式。

![示例](../../images/worksheetfunctionschainingResult.JPG)


下面的示例将对上方的表应用 vlookup 函数，查找 Wrench 在 11 月的销售个数。
```js
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
        var unitSoldInNov = ctx.workbook.functions.vlookup("Wrench", range, 2, false);
        unitSoldInNov.load();
        return ctx.sync()
        .then(function () {
            console.log(unitSoldInNov.value);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```




下面的示例将使用 vlookup 函数首先找到 Wrench 在 11 月和 12 月的各自销售个数。然后应用 sum 函数来获取这两个月内售出个数的总和。请注意只需加载最终结果，在应用最后的公式时，将对所有中间结果进行计算和使用。

```js
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
        var sumOfTwoLookups = ctx.workbook.functions.sum(
            ctx.workbook.functions.vlookup("Wrench", range, 2, false), 
            ctx.workbook.functions.vlookup("Wrench", range, 3, false)
            );
        sumOfTwoLookups.load();
        return ctx.sync()
        .then(function () {
            console.log(sumOfTwoLookups.value);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });


```

## [受支持的工作表函数的列表](#受支持的工作表函数的列表)

| 方法           | 返回类型    |说明|注释 |
|:---------------|:--------|:----------|:-----|
|[ABS 函数](https://support.office.com/en-us/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c)| FunctionResult |返回数字的绝对值|
|[ACCRINT 函数](https://support.office.com/en-us/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74)| FunctionResult |返回定期支付利息的债券的应计利息|
|[ACCRINTM 函数](https://support.office.com/en-us/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7)| FunctionResult |返回在到期日支付利息的债券的应计利息|
|[ACOS 函数](https://support.office.com/en-us/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b)| FunctionResult |返回一个数的反余弦值|
|[ACOSH 函数](https://support.office.com/en-us/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe)| FunctionResult |返回一个数的反双曲余弦值|
|[ACOT 函数](https://support.office.com/en-us/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905)| FunctionResult |返回一个数的反余切值|
|[ACOTH 函数](https://support.office.com/en-us/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f)| FunctionResult |返回一个数的双曲反余切值|
|[AMORDEGRC 函数](https://support.office.com/en-us/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51)| FunctionResult |通过使用折旧系数，返回每个会计期间的折旧值|
|[AMORLINC 函数](https://support.office.com/en-us/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8)| FunctionResult |返回每个会计期间的折旧值|
|[AND 函数](https://support.office.com/en-us/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9)| FunctionResult |如果其所有参数都为 TRUE，则返回 TRUE|
|[ARABIC 函数](https://support.office.com/en-us/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f)| FunctionResult |将罗马数字转换为阿拉伯数字|
|[AREAS 函数](https://support.office.com/en-us/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152)| FunctionResult |返回引用中包含的区域个数|
|[ASC 函数](https://support.office.com/en-us/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266)| FunctionResult |将字符串中的全角（双字节）英文字母或片假名更改为半角（单字节）字符|
|[ASIN 函数](https://support.office.com/en-us/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347)| FunctionResult |返回一个数的反正弦值|
|[ASINH 函数](https://support.office.com/en-us/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c)| FunctionResult |返回一个数的反双曲正弦值|
|[ATAN 函数](https://support.office.com/en-us/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543)| FunctionResult |返回一个数的反正切值|
|[ATAN2 函数](https://support.office.com/en-us/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033)| FunctionResult |返回从 x 坐标和 y 坐标的反正切值|
|[ATANH 函数](https://support.office.com/en-us/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90)| FunctionResult |返回某一数字的反双曲正切值|
|[AVEDEV 函数](https://support.office.com/en-us/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639)| FunctionResult |返回一组数据点到其算术平均值的绝对偏差的平均值|
|[AVERAGE 函数](https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6)| FunctionResult |返回其参数的平均值|
|[AVERAGEA 函数](https://support.office.com/en-us/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091)| FunctionResult |返回其参数的平均值，包括数字、文本和逻辑值|
|[AVERAGEIF 函数](https://support.office.com/en-us/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642)| FunctionResult |返回区域内满足给定条件的所有单元格的平均值（算术平均值）|
|[AVERAGEIFS 函数](https://support.office.com/en-us/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690)| FunctionResult |返回满足多个条件的所有单元格的平均值（算术平均值）|
|[BAHTTEXT 函数](https://support.office.com/en-us/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c)| FunctionResult |使用 ß（铢）货币格式将数字转换为文本|
|[BASE 函数](https://support.office.com/en-us/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342)| FunctionResult |将数字转换成具有给定基数的文本表示形式|
|[BESSELI 函数](https://support.office.com/en-us/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df)| FunctionResult |返回修正的贝塞耳函数 In(x)|
|[BESSELJ 函数](https://support.office.com/en-us/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7)| FunctionResult |返回贝塞耳函数 Jn(x)|
|[BESSELK 函数](https://support.office.com/en-us/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70)| FunctionResult |返回修正的贝塞耳函数 Kn(x)|
|[BESSELY 函数](https://support.office.com/en-us/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2)| FunctionResult |返回贝赛耳函数 Yn(x)|
|[BETA.DIST 函数](https://support.office.com/en-us/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31)| FunctionResult |返回 beta 累积分布函数|
|[BETA.INV 函数](https://support.office.com/en-us/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb)| FunctionResult |返回指定的 beta 分布累积分布函数的反函数|
|[BIN2DEC 函数](https://support.office.com/en-us/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c)| FunctionResult |将二进制数转换为十进制|
|[BIN2HEX 函数](https://support.office.com/en-us/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc)| FunctionResult |将二进制数转换为十六进制|
|[BIN2OCT 函数](https://support.office.com/en-us/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd)| FunctionResult |将二进制数转换为八进制|
|[BINOM.DIST 函数](https://support.office.com/en-us/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c)| FunctionResult |返回一元二项式分布的概率|
|[BINOM.DIST.RANGE 函数](https://support.office.com/en-us/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595)| FunctionResult |返回使用二项式分布的试验结果的概率|
|[BINOM.INV 函数](https://support.office.com/en-us/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9)| FunctionResult |返回一个数值，它是使得累积二项式分布的函数值小于或等于临界值的最小整数|
|[BITAND 函数](https://support.office.com/en-us/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a)| FunctionResult |返回两个数字的“按位与”|
|[BITLSHIFT 函数](https://support.office.com/en-us/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c)| FunctionResult |返回按照 shift_amount 位数左移后得到的数值|
|[BITOR 函数](https://support.office.com/en-us/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2)| FunctionResult |返回 2 个数字的按位“或”|
|[BITRSHIFT 函数](https://support.office.com/en-us/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c)| FunctionResult |返回按照 shift_amount 位数右移后得到的数值|
|[BITXOR 函数](https://support.office.com/en-us/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4)| FunctionResult |返回两个数字的按位“异或”值|
|[CEILING.MATH 函数](https://support.office.com/en-us/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8)| FunctionResult |将数值向上舍入为最接近的整数或最接近的基数的倍数|
|[CEILING.PRECISE 函数](https://support.office.com/en-us/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb)| FunctionResult |将数值四舍五入到最接近的整数或最接近的基数的倍数。不论数字是否带有符号，都将数字向上舍入。|
|[CHAR 函数](https://support.office.com/en-us/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a)| FunctionResult |返回由代码数字指定的字符|
|[CHISQ.DIST 函数](https://support.office.com/en-us/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732)| FunctionResult |返回累积 beta 分布的概率密度函数|
|[CHISQ.DIST.RT 函数](https://support.office.com/en-us/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2)| FunctionResult |返回 χ2 分布的收尾概率|
|[CHISQ.INV 函数](https://support.office.com/en-us/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f)| FunctionResult |返回累积 beta 分布的概率密度函数|
|[CHISQ.INV.RT 函数](https://support.office.com/en-us/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe)| FunctionResult |返回 χ2 分布的收尾概率的反函数|
|[CHOOSE 函数](https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc)| FunctionResult |从值列表中选择一个值|
|[CLEAN 函数](https://support.office.com/en-us/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41)| FunctionResult |删除文本中的所有非打印字符|
|[CODE 函数](https://support.office.com/en-us/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928)| FunctionResult |返回文本字符串中第一个字符的数字代码|
|[COLUMNS 函数](https://support.office.com/en-us/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca)| FunctionResult |返回引用中的列数|
|[COMBIN 函数](https://support.office.com/en-us/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a)| FunctionResult |返回给定数目对象的组合数|
|[COMBINA 函数](https://support.office.com/en-us/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d)| FunctionResult |返回给定项数的组合数（包含重复项）|
|[COMPLEX 函数](https://support.office.com/en-us/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128)| FunctionResult |将实部系数和虚部系数转换为复数|
|[CONCATENATE 函数](https://support.office.com/en-us/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d)| FunctionResult |将几个文本项合并为一个文本项|
|[CONFIDENCE.NORM 函数](https://support.office.com/en-us/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4)| FunctionResult |返回总体平均数的置信区间|
|[CONFIDENCE.T 函数](https://support.office.com/en-us/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53)| FunctionResult |使用学生 t 分布返回总体平均数的置信区间|
|[CONVERT 函数](https://support.office.com/en-us/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2)| FunctionResult |将数字从一种度量体系转换为另一种度量体系|
|[COS 函数](https://support.office.com/en-us/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05)| FunctionResult |返回一个数的余弦值|
|[COSH 函数](https://support.office.com/en-us/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555)| FunctionResult |返回一个数字的双曲余弦值|
|[COT 函数](https://support.office.com/en-us/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a)| FunctionResult |返回一个角度的余切值|
|[COTH 函数](https://support.office.com/en-us/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5)| FunctionResult |返回一个数字的双曲余切值|
|[COUNT 函数](https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c)| FunctionResult |计算参数表中的数字个数|
|[COUNTA 函数](https://support.office.com/en-us/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509)| FunctionResult |计算参数列表中值的数量|
|[COUNTBLANK 函数](https://support.office.com/en-us/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22)| FunctionResult |计算在一定范围内的空单元格数量|
|[COUNTIF 函数](https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34)| FunctionResult |计算某个区域中满足给定条件的单元格数目|
|[COUNTIFS 函数](https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842)| FunctionResult |计算某个区域中满足多个条件的单元格数目|
|[COUPDAYBS 函数](https://support.office.com/en-us/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872)| FunctionResult |返回从票息期开始到结算日之间的天数|
|[COUPDAYS 函数](https://support.office.com/en-us/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671)| FunctionResult |返回包含结算日的票息期的天数|
|[COUPDAYSNC 函数](https://support.office.com/en-us/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547)| FunctionResult |返回从结算日到下一票息支付日之间的天数|
|[COUPNCD 函数](https://support.office.com/en-us/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f)| FunctionResult |返回结算日后的下一票息支付日|
|[COUPNUM 函数](https://support.office.com/en-us/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522)| FunctionResult |返回结算日与到期日之间可支付的票息数|
|[COUPPCD 函数](https://support.office.com/en-us/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3)| FunctionResult |返回结算日前的上一票息支付日|
|[CSC 函数](https://support.office.com/en-us/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1)| FunctionResult |返回一个角度的余割值|
|[CSCH 函数](https://support.office.com/en-us/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb)| FunctionResult |返回一个角度的双曲余割值|
|[CUMIPMT 函数](https://support.office.com/en-us/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606)| FunctionResult |返回两个付款期之间为贷款累积支付的利息|
|[CUMPRINC 函数](https://support.office.com/en-us/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d)| FunctionResult |返回两个付款期之间为贷款累积支付的本金|
|[DATE 函数](https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349)| FunctionResult |返回特定日期的序列号|
|[DATEVALUE 函数](https://support.office.com/en-us/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252)| FunctionResult |将以文本表达的日期转换为序列号|
|[DAVERAGE 函数](https://support.office.com/en-us/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee)| FunctionResult |返回所选数据库条目的平均值|
|[DAY 函数](https://support.office.com/en-us/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101)| FunctionResult |将序列号转换为月份中的某一天|
|[DAYS 函数](https://support.office.com/en-us/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df)| FunctionResult |返回两个日期间相差的天数|
|[DAYS360 函数](https://support.office.com/en-us/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a)| FunctionResult |按每年 360 天计算两个日期间相差的天数|
|[DB 函数](https://support.office.com/en-us/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7)| FunctionResult |使用固定余额递减法返回指定周期内某项资产的折旧值|
|[DBCS 函数](https://support.office.com/en-us/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5)| FunctionResult |将字符串中的半角（单字节）英文字母或片假名更改为全角（双字节）字符|
|[DCOUNT 函数](https://support.office.com/en-us/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1)| FunctionResult |计算数据库中包含数字的单元格数量|
|[DCOUNTA 函数](https://support.office.com/en-us/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244)| FunctionResult |计算数据库中的非空单元格的数量|
|[DDB 函数](https://support.office.com/en-us/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5)| FunctionResult |使用双倍余额递减法或其他指定方法返回某项资产在指定周期内的折旧值|
|[DEC2BIN 函数](https://support.office.com/en-us/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838)| FunctionResult |将十进制数转换为二进制|
|[DEC2HEX 函数](https://support.office.com/en-us/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619)| FunctionResult |将十进制数转换为十六进制|
|[DEC2OCT 函数](https://support.office.com/en-us/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f)| FunctionResult |将十进制数转换为八进制|
|[DECIMAL 函数](https://support.office.com/en-us/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e)| FunctionResult |按给定基数将数字的文本表示形式转换成十进制数|
|[DEGREES 函数](https://support.office.com/en-us/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1)| FunctionResult |将弧度转换为角度|
|[DELTA 函数](https://support.office.com/en-us/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432)| FunctionResult |测试两个值是否相等|
|[DEVSQ 函数](https://support.office.com/en-us/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444)| FunctionResult |返回偏差平方和|
|[DGET 函数](https://support.office.com/en-us/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e)| FunctionResult |从数据库中提取符合指定条件的单个记录|
|[DISC 函数](https://support.office.com/en-us/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53)| FunctionResult |返回债券的贴现率|
|[DMAX 函数](https://support.office.com/en-us/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2)| FunctionResult |返回所选数据库条目中的最大值|
|[DMIN 函数](https://support.office.com/en-us/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3)| FunctionResult |返回所选数据库条目中的最小值|
|[DOLLAR 函数](https://support.office.com/en-us/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611)| FunctionResult |使用 $（美元）货币格式将数字转换为文本|
|[DOLLARDE 函数](https://support.office.com/en-us/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427)| FunctionResult |将以分数表示的货币值转换为以小数表示的货币值|
|[DOLLARFR 函数](https://support.office.com/en-us/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495)| FunctionResult |将以小数表示的货币值转换为以分数表示的货币值|
|[DPRODUCT 函数](https://support.office.com/en-us/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31)| FunctionResult |将与数据库中的条件匹配的记录的特定字段中的值相乘|
|[DSTDEV 函数](https://support.office.com/en-us/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96)| FunctionResult |根据所选数据库条目中的样本估算数据的标准偏差|
|[DSTDEVP 函数](https://support.office.com/en-us/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b)| FunctionResult |以数据库选定项作为样本总体，计算数据的标准偏差|
|[DSUM 函数](https://support.office.com/en-us/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b)| FunctionResult |将数据库中与条件匹配的记录字段列中的数字进行求和|
|[DURATION 函数](https://support.office.com/en-us/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038)| FunctionResult |返回定期支付利息的债券的年持续时间|
|[DVAR 函数](https://support.office.com/en-us/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1)| FunctionResult |根据所选数据库条目中的样本估算数据的方差|
|[DVARP 函数](https://support.office.com/en-us/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc)| FunctionResult |以数据库选定项作为样本总体，计算数据的总体方差|
|[EDATE 函数](https://support.office.com/en-us/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5)| FunctionResult |返回一串日期，指示起始日期之前/之后的月数|
|[EFFECT 函数](https://support.office.com/en-us/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4)| FunctionResult |返回年有效利率|
|[EOMONTH 函数](https://support.office.com/en-us/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628)| FunctionResult |返回一串日期，表示指定月数之前或之后的月份的最后一天|
|[ERF 函数](https://support.office.com/en-us/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349)| FunctionResult |返回误差函数|
|[ERF.PRECISE 函数](https://support.office.com/en-us/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0)| FunctionResult |返回误差函数|
|[ERFC 函数](https://support.office.com/en-us/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3)| FunctionResult |返回补余误差函数|
|[ERFC.PRECISE 函数](https://support.office.com/en-us/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273)| FunctionResult |返回在 x 和无穷大之间集成的补余 ERF 函数|
|[ERROR.TYPE 函数](https://support.office.com/en-us/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa)| FunctionResult |返回对应于一种错误类型的数字|
|[EVEN 函数](https://support.office.com/en-us/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9)| FunctionResult |将数字向上舍入到最近的偶数|
|[EXACT 函数](https://support.office.com/en-us/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926)| FunctionResult |检查两个文本值是否相同|
|[EXP 函数](https://support.office.com/en-us/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe)| FunctionResult |返回 e 的 n 次方|
|[EXPON.DIST 函数](https://support.office.com/en-us/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e)| FunctionResult |返回指数分布|
|[F.DIST 函数](https://support.office.com/en-us/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d)| FunctionResult |返回 F 概率分布|
|[F.DIST.RT 函数](https://support.office.com/en-us/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520)| FunctionResult |返回 F 概率分布|
|[F.INV 函数](https://support.office.com/en-us/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe)| FunctionResult |返回 F 概率分布的逆函数值|
|[F.INV.RT 函数](https://support.office.com/en-us/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00)| FunctionResult |返回 F 概率分布的逆函数值|
|[FACT 函数](https://support.office.com/en-us/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3)| FunctionResult |返回某数的阶乘|
|[FACTDOUBLE 函数](https://support.office.com/en-us/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8)| FunctionResult |返回数字的双阶乘|
|[FALSE 函数](https://support.office.com/en-us/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904)| FunctionResult |返回逻辑值 FALSE|
|[FIND 函数和 FINDB 函数](https://support.office.com/en-us/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628)| FunctionResult |在一个文本值中查找另一个（区分大小写）|
|[FISHER 函数](https://support.office.com/en-us/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69)| FunctionResult |返回 Fisher 变换值|
|[FISHERINV 函数](https://support.office.com/en-us/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb)| FunctionResult |返回 Fisher 逆变换值|
|[FIXED 函数](https://support.office.com/en-us/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a)| FunctionResult |将数字格式化为具有固定数量的小数的文本|
|[FLOOR 函数](https://support.office.com/en-us/article/FLOOR-function-14bb497c-24f2-4e04-b327-b0b4de5a8886)| FunctionResult |将数字向零的方向向下舍入|
|[FLOOR.MATH 函数](https://support.office.com/en-us/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5)| FunctionResult |将数字向下舍入为最接近的整数或最接近的基数的倍数|
|[FLOOR.PRECISE 函数](https://support.office.com/en-us/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e)| FunctionResult |将数字向下舍入为最接近的整数或最接近的基数的倍数。不论数字是否带有符号，都将数字向下舍入。|
|[FORECAST 函数](https://support.office.com/en-us/article/FORECAST-function-50ca49c9-7b40-4892-94e4-7ad38bbeda99)| FunctionResult |返回一个值和线性趋势|
|[FORECAST.ETS 函数](https://support.office.com/en-us/article/FORECASTETS-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |通过使用指数平滑 (ETS) 算法的 AAA 版本返回基于现有的（历史）值得出的未来值|
|[FORECAST.ETS.CONFINT 函数](https://support.office.com/en-us/article/FORECASTETSCONFINT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |返回指定目标日期预测值的置信区间|
|[FORECAST.ETS.SEASONALITY 函数](https://support.office.com/en-us/article/FORECASTETSSEASONALITY-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |返回在一系列指定时间里 Excel 检测的重复模式的长度|
|[FORECAST.ETS.STAT 函数](https://support.office.com/en-us/article/FORECASTETSSTAT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |返回作为时间序列预报结果的统计值|
|[FORECAST.LINEAR 函数](https://support.office.com/en-us/article/FORECASTLINEAR-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |返回基于现有值的未来值|
|[FV 函数](https://support.office.com/en-us/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3)| FunctionResult |返回一项投资的未来值|
|[FVSCHEDULE 函数](https://support.office.com/en-us/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d)| FunctionResult |返回在应用一系列复利后，初始本金的终值|
|[GAMMA 函数](https://support.office.com/en-us/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297)| FunctionResult |返回 Gamma 函数值|
|[GAMMA.DIST 函数](https://support.office.com/en-us/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def)| FunctionResult |返回 γ 分布|
|[GAMMA.INV 函数](https://support.office.com/en-us/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18)| FunctionResult |返回 γ 累积分布的反函数|
|[GAMMALN 函数](https://support.office.com/en-us/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9)| FunctionResult |返回 γ 函数的自然对数 Γ(x)|
|[GAMMALN.PRECISE 函数](https://support.office.com/en-us/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599)| FunctionResult |返回 γ 函数的自然对数 Γ(x)|
|[GAUSS 函数](https://support.office.com/en-us/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33)| FunctionResult |返回比标准正态累积分布小 0.5 的值|
|[GCD 函数](https://support.office.com/en-us/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a)| FunctionResult |返回最大公约数|
|[GEOMEAN 函数](https://support.office.com/en-us/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5)| FunctionResult |返回几何平均数|
|[GESTEP 函数](https://support.office.com/en-us/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df)| FunctionResult |测试某个数字是否大于阈值|
|[HARMEAN 函数](https://support.office.com/en-us/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6)| FunctionResult |返回调和平均值|
|[HEX2BIN 函数](https://support.office.com/en-us/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1)| FunctionResult |将十六进制数转换为二进制|
|[HEX2DEC 函数](https://support.office.com/en-us/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e)| FunctionResult |将十六进制数转换为十进制|
|[HEX2OCT 函数](https://support.office.com/en-us/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912)| FunctionResult |将十六进制数转换为八进制|
|[HLOOKUP 函数](https://support.office.com/en-us/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f)| FunctionResult |在数组的顶行中查找并返回指定单元格的值|
|[HOUR 函数](https://support.office.com/en-us/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7)| FunctionResult |将序列号转换为小时|
|[HYPERLINK 函数](https://support.office.com/en-us/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f)| FunctionResult |创建一个快捷方式或链接，以便打开一个存储在网络服务器、内部网或 Internet 上的文档|
|[HYPGEOM.DIST 函数](https://support.office.com/en-us/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf)| FunctionResult |返回超几何分布|
|[IF 函数](https://support.office.com/en-us/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)| FunctionResult |指定要执行的逻辑测试|
|[IMABS 函数](https://support.office.com/en-us/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1)| FunctionResult |返回复数的绝对值（模数）|
|[IMAGINARY 函数](https://support.office.com/en-us/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a)| FunctionResult |返回复数的虚部系数|
|[IMARGUMENT 函数](https://support.office.com/en-us/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a)| FunctionResult |返回以弧度表示的角 - 参数 θ|
|[IMCONJUGATE 函数](https://support.office.com/en-us/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42)| FunctionResult |返回复数的共轭复数|
|[IMCOS 函数](https://support.office.com/en-us/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c)| FunctionResult |返回复数的余弦值|
|[IMCOSH 函数](https://support.office.com/en-us/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff)| FunctionResult |返回复数的双曲余弦值|
|[IMCOT 函数](https://support.office.com/en-us/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c)| FunctionResult |返回复数的余切值|
|[IMCSC 函数](https://support.office.com/en-us/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323)| FunctionResult |返回复数的余割值|
|[IMCSCH 函数](https://support.office.com/en-us/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9)| FunctionResult |返回复数的双曲余割值|
|[IMDIV 函数](https://support.office.com/en-us/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f)| FunctionResult |返回两个复数之商|
|[IMEXP 函数](https://support.office.com/en-us/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f)| FunctionResult |返回复数的指数值|
|[IMLN 函数](https://support.office.com/en-us/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8)| FunctionResult |返回复数的自然对数|
|[IMLOG10 函数](https://support.office.com/en-us/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5)| FunctionResult |返回以 10 为底的复数的对数|
|[IMLOG2 函数](https://support.office.com/en-us/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51)| FunctionResult |返回以 2 为底的复数的对数|
|[IMPOWER 函数](https://support.office.com/en-us/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39)| FunctionResult |返回复数的整数幂|
|[IMPRODUCT 函数](https://support.office.com/en-us/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba)| FunctionResult |返回从 2 到 255 个复数的乘积|
|[IMREAL 函数](https://support.office.com/en-us/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366)| FunctionResult |返回复数的实部系数|
|[IMSEC 函数](https://support.office.com/en-us/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0)| FunctionResult |返回复数的正割值|
|[IMSECH 函数](https://support.office.com/en-us/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b)| FunctionResult |返回复数的双曲正割值|
|[IMSIN 函数](https://support.office.com/en-us/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6)| FunctionResult |返回复数的正弦值|
|[IMSINH 函数](https://support.office.com/en-us/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d)| FunctionResult |返回复数的双曲正弦值|
|[IMSQRT 函数](https://support.office.com/en-us/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70)| FunctionResult |返回复数的平方根|
|[IMSUB 函数](https://support.office.com/en-us/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054)| FunctionResult |返回两个复数的差值|
|[IMSUM 函数](https://support.office.com/en-us/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f)| FunctionResult |返回复数的和|
|[IMTAN 函数](https://support.office.com/en-us/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132)| FunctionResult |返回复数的正切值|
|[INT 函数](https://support.office.com/en-us/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef)| FunctionResult |将数值向下舍入到最接近的整数|
|[INTRATE 函数](https://support.office.com/en-us/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f)| FunctionResult |返回完全投资型债券的利率|
|[IPMT 函数](https://support.office.com/en-us/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f)| FunctionResult |返回给定期间内投资所支付的利息|
|[IRR 函数](https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc)| FunctionResult |返回一系列现金流的内部收益率|
|[ISERR 函数](https://support.office.com/en-us/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果数值是除 #N/A 之外的错误值，则返回 TRUE|
|[ISERROR 函数](https://support.office.com/en-us/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果数值是任何错误值，则一律返回 TRUE|
|[ISEVEN 函数](https://support.office.com/en-us/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356)| FunctionResult |如果数值为偶数，则返回 TRUE|
|[ISFORMULA 函数](https://support.office.com/en-us/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5)| FunctionResult |如果存在对包含公式的单元格的引用，则返回 TRUE|
|[ISLOGICAL 函数](https://support.office.com/en-us/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果值为逻辑值，则返回 TRUE|
|[ISNA 函数](https://support.office.com/en-us/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果该值为 #N/A 错误，则返回 TRUE|
|[ISNONTEXT 函数](https://support.office.com/en-us/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果值不是文本，则返回 TRUE|
|[ISNUMBER 函数](https://support.office.com/en-us/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果值为数字，则返回 TRUE|
|[ISO.CEILING 函数](https://support.office.com/en-us/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83)| FunctionResult |将数字向上舍入到最接近的整数或最接近的基数的倍数|
|[ISODD 函数](https://support.office.com/en-us/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果数字为奇数，则返回 TRUE|
|[ISOWEEKNUM 函数](https://support.office.com/en-us/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e)| FunctionResult |返回一年中给定日期的 ISO 周数的数目|
|[ISPMT 函数](https://support.office.com/en-us/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc)| FunctionResult |计算指定的投资期间支付的利息|
|[ISREF 函数](https://support.office.com/en-us/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果值为引用值，则返回 TRUE|
|[ISTEXT 函数](https://support.office.com/en-us/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |如果值为文本，则返回 TRUE|
|[KURT 函数](https://support.office.com/en-us/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab)| FunctionResult |返回一组数据的峰值|
|[LARGE 函数](https://support.office.com/en-us/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64)| FunctionResult |返回数据集中第 k 个最大值|
|[LCM 函数](https://support.office.com/en-us/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94)| FunctionResult |返回最小公倍数|
|[LEFT、LEFTB 函数](https://support.office.com/en-us/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c)| FunctionResult |返回一个文本值的最左端字符|
|[LEN、LENB 函数](https://support.office.com/en-us/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb)| FunctionResult |返回文本字符串中的字符数|
|[LN 函数](https://support.office.com/en-us/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f)| FunctionResult |返回数值的自然对数|
|[LOG 函数](https://support.office.com/en-us/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280)| FunctionResult |返回一个数在指定底下的对数|
|[LOG10 函数](https://support.office.com/en-us/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211)| FunctionResult |返回以 10 为底的对数|
|[LOGNORM.DIST 函数](https://support.office.com/en-us/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070)| FunctionResult |返回对数正态分布|
|[LOGNORM.INV 函数](https://support.office.com/en-us/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600)| FunctionResult |返回对数正态分布的反函数|
|[LOOKUP 函数](https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb)| FunctionResult |在向量或数组中查找值|
|[LOWER 函数](https://support.office.com/en-us/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4)| FunctionResult |将文本转换为小写|
|[MATCH 函数](https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a)| FunctionResult |在引用或数组中查找值|
|[MAX 函数](https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098)| FunctionResult |返回参数列表中的最大值|
|[MAXA 函数](https://support.office.com/en-us/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d)| FunctionResult |返回参数列表中的最大值，包括数字、文本和逻辑值|
|[MDURATION 函数](https://support.office.com/en-us/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c)| FunctionResult |为假定票面值为 100 元的债券返回麦考利修正持续时间|
|[MEDIAN 函数](https://support.office.com/en-us/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2)| FunctionResult |返回给定数字的中值|
|[MID、MIDB 函数](https://support.office.com/en-us/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028)| FunctionResult |从指定位置开始，返回文本字符串中特定数量的字符。|
|[MIN 函数](https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152)| FunctionResult |返回参数列表中的最小值|
|[MINA 函数](https://support.office.com/en-us/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3)| FunctionResult |返回参数列表中的最小值，包括数字、文本和逻辑值|
|[MINUTE 函数](https://support.office.com/en-us/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589)| FunctionResult |将序列号转换为分钟|
|[MIRR 函数](https://support.office.com/en-us/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524)| FunctionResult |返回内部收益率，它的正现金流和负现金流以不同的比率融资|
|[MOD 函数](https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3)| FunctionResult |返回除法的余数|
|[MONTH 函数](https://support.office.com/en-us/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8)| FunctionResult |将序列号转换为月|
|[MROUND 函数](https://support.office.com/en-us/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427)| FunctionResult |返回舍入到所需倍数的数值|
|[MULTINOMIAL 函数](https://support.office.com/en-us/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6)| FunctionResult |返回一组数字的多项式|
|[N 函数](https://support.office.com/en-us/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9)| FunctionResult |返回转换为数字的值|
|[NA 函数](https://support.office.com/en-us/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c)| FunctionResult |返回错误值 #N/A|
|[NEGBINOM.DIST 函数](https://support.office.com/en-us/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599)| FunctionResult |返回负二项式分布函数值|
|[NETWORKDAYS 函数](https://support.office.com/en-us/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7)| FunctionResult |返回两个日期之间的完整工作日数|
|[NETWORKDAYS.INTL 函数](https://support.office.com/en-us/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28)| FunctionResult |使用能够指示哪些以及有多少天是周末的参数返回两个日期之间的完整工作日数|
|[NOMINAL 函数](https://support.office.com/en-us/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b)| FunctionResult |返回年度的单利|
|[NORM.DIST 函数](https://support.office.com/en-us/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d)| FunctionResult |返回正态分布函数值|
|[NORM.INV 函数](https://support.office.com/en-us/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13)| FunctionResult |返回正态分布的反函数|
|[NORM.S.DIST 函数](https://support.office.com/en-us/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88)| FunctionResult |返回标准正态分布函数值|
|[NORM.S.INV 函数](https://support.office.com/en-us/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1)| FunctionResult |返回标准正态分布的反函数|
|[NOT 函数](https://support.office.com/en-us/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77)| FunctionResult |反转其参数的逻辑|
|[Now 函数](https://support.office.com/en-us/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46)| FunctionResult |返回当前日期和时间的序列号|
|[NPER 函数](https://support.office.com/en-us/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815)| FunctionResult |返回一项投资的周期数量|
|[NPV 函数](https://support.office.com/en-us/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568)| FunctionResult |基于一系列定期现金流和贴现率返回投资的净现值|
|[NUMBERVALUE 函数](https://support.office.com/en-us/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879)| FunctionResult |按独立于区域设置的方式将文本转换为数字|
|[OCT2BIN 函数](https://support.office.com/en-us/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589)| FunctionResult |将八进制数转换为二进制|
|[OCT2DEC 函数](https://support.office.com/en-us/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554)| FunctionResult |将八进制数转换为十进制|
|[OCT2HEX 函数](https://support.office.com/en-us/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f)| FunctionResult |将八进制数转换为十六进制|
|[ODD 函数](https://support.office.com/en-us/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98)| FunctionResult |将数值向上舍入到最接近的奇数|
|[ODDFPRICE 函数](https://support.office.com/en-us/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1)| FunctionResult |返回每张票面为 100 元且第一期为奇数的债券的现价|
|[ODDFYIELD 函数](https://support.office.com/en-us/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37)| FunctionResult |返回第一期为奇数的债券的收益|
|[ODDLPRICE 函数](https://support.office.com/en-us/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4)| FunctionResult |返回每张票面为 100 元且最后一期为奇数的债券的现价|
|[ODDLYIELD 函数](https://support.office.com/en-us/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238)| FunctionResult |返回最后一期为奇数的债券的收益|
|[OR 函数](https://support.office.com/en-us/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0)| FunctionResult |如果任意参数为 TRUE，则返回 TRUE|
|[PDURATION 函数](https://support.office.com/en-us/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf)| FunctionResult |返回投资达到指定的值所需的期数|
|[PERCENTILE.EXC 函数](https://support.office.com/en-us/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba)| FunctionResult |返回数组的 K 百分点值，K 介于 0 与 1 之间，不含 0 与 1|
|[PERCENTILE.INC 函数](https://support.office.com/en-us/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed)| FunctionResult |返回数组的 K 百分点值|
|[PERCENTRANK.EXC 函数](https://support.office.com/en-us/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314)| FunctionResult |返回特定数值在一个数据集中的百分比排名（介于 0 与 1 之间，不含 0 与 1）|
|[PERCENTRANK.INC 函数](https://support.office.com/en-us/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a)| FunctionResult |返回一组数据中的值的百分比排名|
|[PERMUT 函数](https://support.office.com/en-us/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3)| FunctionResult |返回给定数目对象的排列数|
|[PERMUTATIONA 函数](https://support.office.com/en-us/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e)| FunctionResult |返回从给定元素数目的集合中选取若干（包括重复项）元素的排列数|
|[PHI 函数](https://support.office.com/en-us/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c)| FunctionResult |返回标准正态分布的密度函数值|
|[PI 函数](https://support.office.com/en-us/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b)| FunctionResult |返回 pi 值|
|[PMT 函数](https://support.office.com/en-us/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441)| FunctionResult |返回年金的定期支付额|
|[POISSON.DIST 函数](https://support.office.com/en-us/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636)| FunctionResult |返回泊松分布|
|[POWER 函数](https://support.office.com/en-us/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a)| FunctionResult |返回某数的乘幂结果|
|[PPMT 函数](https://support.office.com/en-us/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b)| FunctionResult |返回对给定期间内的投资所支付的本金|
|[PRICE 函数](https://support.office.com/en-us/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a)| FunctionResult |返回每张票面为 100 元且定期支付利息的债券的现价|
|[PRICEDISC 函数](https://support.office.com/en-us/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3)| FunctionResult |返回每张票面为 100 元的已贴现债券的现价|
|[PRICEMAT 函数](https://support.office.com/en-us/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77)| FunctionResult |返回每张票面为 100 元且在到期日支付利息的债券的现价|
|[PROB 函数](https://support.office.com/en-us/article/PROB-function-9ac30561-c81c-4259-8253-34f0a238fc49)| FunctionResult |返回这些值存在于两个限定值之间的范围中的概率|
|[PRODUCT 函数](https://support.office.com/en-us/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce)| FunctionResult |将其参数相乘|
|[PROPER 函数](https://support.office.com/en-us/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94)| FunctionResult |使一个文本值的每个词的首字母大写|
|[PV 函数](https://support.office.com/en-us/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd)| FunctionResult |返回一项投资的当前值|
|[QUARTILE.EXC 函数](https://support.office.com/en-us/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad)| FunctionResult |基于从 0 到 1 之间（不含 0 与 1）的百分点值，返回一组数据的四分位点|
|[QUARTILE.INC 函数](https://support.office.com/en-us/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d)| FunctionResult |返回一组数据的四分位点|
|[QUOTIENT 函数](https://support.office.com/en-us/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee)| FunctionResult |返回除法结果的整数部分|
|[RADIANS 函数](https://support.office.com/en-us/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf)| FunctionResult |将度转换为弧度|
|[RAND 函数](https://support.office.com/en-us/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73)| FunctionResult |返回 0 和 1 之间的一个随机数|
|[RANDBETWEEN 函数](https://support.office.com/en-us/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685)| FunctionResult |返回指定数字之间的随机数|
|[RANK.AVG 函数](https://support.office.com/en-us/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a)| FunctionResult |返回某数字在一列数字中的排名|
|[RANK.EQ 函数](https://support.office.com/en-us/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40)| FunctionResult |返回某数字在一列数字中的排名|
|[RATE 函数](https://support.office.com/en-us/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce)| FunctionResult |返回年金的定期利率|
|[RECEIVED 函数](https://support.office.com/en-us/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5)| FunctionResult |返回完全投资型债券到期收回的金额|
|[REPLACE、REPLACEB 函数](https://support.office.com/en-us/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a)| FunctionResult |替换文本中的字符|
|[REPT 函数](https://support.office.com/en-us/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061)| FunctionResult |以给定的次数重复文本|
|[RIGHT、RIGHTB 函数](https://support.office.com/en-us/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f)| FunctionResult |返回一个文本值的最右端字符|
|[ROMAN 函数](https://support.office.com/en-us/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5)| FunctionResult |将阿拉伯数字转换为文本形式的罗马数字|
|[ROUND 函数](https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c)| FunctionResult |将数字舍入到指定位数|
|[ROUNDDOWN 函数](https://support.office.com/en-us/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53)| FunctionResult |将数字向零的方向向下舍入|
|[ROUNDUP 函数](https://support.office.com/en-us/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7)| FunctionResult |将数字向远离零的方向向上舍入|
|[ROWS 函数](https://support.office.com/en-us/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597)| FunctionResult |返回引用中的行数|
|[RRI 函数](https://support.office.com/en-us/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4)| FunctionResult |返回某项投资增长的等效利率|
|[RTD 函数](https://support.office.com/en-us/article/RTD-function-e0cc001a-56f0-470a-9b19-9455dc0eb593)| FunctionResult |从一个支持 COM 自动化的程序中获取实时数据|
|[SEC 函数](https://support.office.com/en-us/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7)| FunctionResult |返回一个角度的正割值|
|[SECH 函数](https://support.office.com/en-us/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f)| FunctionResult |返回一个角度的双曲正割值|
|[SECOND 函数](https://support.office.com/en-us/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1)| FunctionResult |将序列号转换为秒|
|[SERIESSUM 函数](https://support.office.com/en-us/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637)| FunctionResult |返回基于以下公式的幂级数之和|
|[SHEET 函数](https://support.office.com/en-us/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24)| FunctionResult |返回引用的工作表的工作表编号|
|[SHEETS 函数](https://support.office.com/en-us/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b)| FunctionResult |返回引用中的工作表数|
|[SIGN 函数](https://support.office.com/en-us/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8)| FunctionResult |返回数值的符号|
|[SIN 函数](https://support.office.com/en-us/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602)| FunctionResult |返回给定角的正弦值|
|[SINH 函数](https://support.office.com/en-us/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7)| FunctionResult |返回某一数字的双曲正弦值|
|[SKEW 函数](https://support.office.com/en-us/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa)| FunctionResult |返回一个分布的不对称度|
|[SKEW.P 函数](https://support.office.com/en-us/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb)| FunctionResult |基于总体返回一个分布的不对称度：用来体现某一分布相对其平均值的不对称程度|
|[SLN 函数](https://support.office.com/en-us/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8)| FunctionResult |返回某项资产一个周期的直线折旧值|
|[SMALL 函数](https://support.office.com/en-us/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07)| FunctionResult |返回数据集中第 k 个最小值|
|[SQRT 函数](https://support.office.com/en-us/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf)| FunctionResult |返回正平方根|
|[SQRTPI 函数](https://support.office.com/en-us/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4)| FunctionResult |返回（数字 * pi）的平方根|
|[STANDARDIZE 函数](https://support.office.com/en-us/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775)| FunctionResult |返回正态分布概率值|
|[STDEV.P 函数](https://support.office.com/en-us/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285)| FunctionResult |基于整个样本总体计算标准偏差|
|[STDEV.S 函数](https://support.office.com/en-us/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23)| FunctionResult |基于样本估计标准偏差|
|[STDEVA 函数](https://support.office.com/en-us/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d)| FunctionResult |基于样本估计标准偏差，包括数字、文本和逻辑值|
|[STDEVPA 函数](https://support.office.com/en-us/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c)| FunctionResult |基于整个样本总体计算标准偏差，包括数字、文本和逻辑值|
|[SUBSTITUTE 函数](https://support.office.com/en-us/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332)| FunctionResult |在文本串中用新文本替换旧文本。|
|[SUBTOTAL 函数](https://support.office.com/en-us/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939)| FunctionResult |返回一个数据列表或数据库的分类汇总|
|[SUM 函数](https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89)| FunctionResult |对参数求和|
|[SUMIF 函数](https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b)| FunctionResult |根据给定的标准，对指定的单元格求和|
|[SUMIFS 函数](https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b)| FunctionResult |对区域中满足多个条件的单元格求和|
|[SUMSQ 函数](https://support.office.com/en-us/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307)| FunctionResult |返回所有参数的平方和|
|[SYD 函数](https://support.office.com/en-us/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27)| FunctionResult |返回某项资产在指定期间的年限总额折旧。|
|[T 函数](https://support.office.com/en-us/article/T-function-fb83aeec-45e7-4924-af95-53e073541228)| FunctionResult |将其参数转换为文本|
|[T.DIST 函数](https://support.office.com/en-us/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2)| FunctionResult |返回学生 t 分布的百分点（概率）|
|[T.DIST.2T 函数](https://support.office.com/en-us/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28)| FunctionResult |返回学生 t 分布的百分点（概率）|
|[T.DIST.RT 函数](https://support.office.com/en-us/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda)| FunctionResult |返回学生的 t 分布|
|[T.INV 函数](https://support.office.com/en-us/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e)| FunctionResult |返回作为概率和自由度函数的学生 t 分布的 t 值|
|[T.INV.2T 函数](https://support.office.com/en-us/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17)| FunctionResult |返回学生 t 分布的反函数|
|[TAN 函数](https://support.office.com/en-us/article/TAN-function-08851a40-179f-4052-b789-d7f699447401)| FunctionResult |返回一个数字的正切值|
|[TANH 函数](https://support.office.com/en-us/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c)| FunctionResult |返回一个数字的双曲正切值|
|[TBILLEQ 函数](https://support.office.com/en-us/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c)| FunctionResult |返回短期国库券的等价债券收益|
|[TBILLPRICE 函数](https://support.office.com/en-us/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2)| FunctionResult |返回每张票面为 100 元的短期国库券的现价|
|[TBILLYIELD 函数](https://support.office.com/en-us/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba)| FunctionResult |返回短期国库券的收益|
|[TEXT 函数](https://support.office.com/en-us/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c)| FunctionResult |设置数字格式并将其转换为文本|
|[TIME 函数](https://support.office.com/en-us/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457)| FunctionResult |返回特定时间的序列号|
|[TIMEVALUE 函数](https://support.office.com/en-us/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645)| FunctionResult |将以文本表达的时间转换为序列号|
|[TODAY 函数](https://support.office.com/en-us/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9)| FunctionResult |返回当前日期的序列号|
|[TRIM 函数](https://support.office.com/en-us/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9)| FunctionResult |从文本中删除空格|
|[TRIMMEAN 函数](https://support.office.com/en-us/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3)| FunctionResult |返回数据集内部的平均值|
|[TRUE 函数](https://support.office.com/en-us/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb)| FunctionResult |返回逻辑值 TRUE|
|[TRUNC 函数](https://support.office.com/en-us/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721)| FunctionResult |将数字截断为整数|
|[TYPE 函数](https://support.office.com/en-us/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899)| FunctionResult |返回一个指示数值数据类型的数字|
|[UNICHAR 函数](https://support.office.com/en-us/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8)| FunctionResult |返回给定数值引用的 Unicode 字符|
|[UNICODE 函数](https://support.office.com/en-us/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f)| FunctionResult |返回与文本的第一个字符相对应的数字（码位）|
|[UPPER 函数](https://support.office.com/en-us/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6)| FunctionResult |将文本转换为大写|
|[VALUE 函数](https://support.office.com/en-us/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2)| FunctionResult |将文本参数转换为数字|
|[VAR.P 函数](https://support.office.com/en-us/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a)| FunctionResult |基于整个样本总体计算方差|
|[VAR.S 函数](https://support.office.com/en-us/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b)| FunctionResult |基于样本估计方差|
|[VARA 函数](https://support.office.com/en-us/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07)| FunctionResult |基于样本估计方差，包括数字、文本和逻辑值|
|[VARPA 函数](https://support.office.com/en-us/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96)| FunctionResult |基于整个样本总体计算方差，包括数字、文本和逻辑值|
|[VDB 函数](https://support.office.com/en-us/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73)| FunctionResult |使用余额递减法返回指定周期或部分周期内某项资产的折旧值|
|[VLOOKUP 函数](https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)| FunctionResult |查找数组的首列并在行间移动以返回单元格的值|
|[WEEKDAY 函数](https://support.office.com/en-us/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a)| FunctionResult |将序列号转换为一周中的某一天|
|[WEEKNUM 函数](https://support.office.com/en-us/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340)| FunctionResult |将序列号转换为代表一年中第几周的数字|
|[WEIBULL.DIST 函数](https://support.office.com/en-us/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db)| FunctionResult |返回 Weibull 分布|
|[WORKDAY 函数](https://support.office.com/en-us/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33)| FunctionResult |返回在指定的若干个工作日之前/之后的日期（一串数字）|
|[WORKDAY.INTL 函数](https://support.office.com/en-us/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d)| FunctionResult |返回在指定的若干个工作日之前/之后的日期（一串数字），其中使用参数来指示哪些以及多少天为周末|
|[XIRR 函数](https://support.office.com/en-us/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d)| FunctionResult |返回一组现金流的内部收益率，这些现金流不一定定期发生|
|[XNPV 函数](https://support.office.com/en-us/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7)| FunctionResult |返回一组现金流的净现值，这些现金流不一定定期发生|
|[XOR 函数](https://support.office.com/en-us/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37)| FunctionResult |返回所有参数的逻辑“异或”值|
|[YEAR 函数](https://support.office.com/en-us/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9)| FunctionResult |将序列号转换为年|
|[YEARFRAC 函数](https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8)| FunctionResult |返回表示 start_date 和 end_date 之间的天数占一年总天数的比值|
|[YIELD 函数](https://support.office.com/en-us/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe)| FunctionResult |返回定期支付利息的债券的收益|
|[YIELDDISC 函数](https://support.office.com/en-us/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7)| FunctionResult |返回已贴现债券的年收益；例如，短期国库券|
|[YIELDMAT 函数](https://support.office.com/en-us/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f)| FunctionResult |返回到期付息的债券的年收益|
|[Z.TEST 函数](https://support.office.com/en-us/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee)| FunctionResult |返回 z 检验的收尾概率值|
