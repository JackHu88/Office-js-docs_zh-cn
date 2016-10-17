
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>使用正则表达式激活规则显示 Outlook 外接程序

你可以指定正则表达式规则以在阅读情况下激活 Outlook 外接程序 - 用户在阅读窗格或检查器中查看邮件或约会时，Outlook 会对正则表达式规则求值，以确定是否应激活你的 Outlook 外接程序。用户在撰写项目时，Outlook 不会对这些规则求值。Outlook 还有其他一些不激活外接程序的情况，例如，项目受信息权限管理 (IRM) 保护或在“垃圾邮件”文件夹中的情况。有关详细信息，请参阅 [Outlook 外接程序的激活规则](../outlook/manifests/activation-rules.md)。

可以指定正则表达式作为加载项 XML 清单中 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 规则或 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 规则的一部分。Outlook 基于客户端计算机上浏览器使用的 JavaScript 解释器的规则对正则表达式求值。对于所有 XML 处理器支持的特殊字符列表，Outlook 同样也支持。下表列出了这些特殊字符。您可以指定相应字符的转义顺序，以在正则表达式中使用这些字符，如下表中所述。



|**字符**|**说明**|**要使用的转义序列**|
|:-----|:-----|:-----|
|"|双引号|&amp;quot;|
|&amp;|与号|&amp;amp;|
|'|撇号|&amp;apos;|
|<|小于号|&amp;lt;|
|>|大于号|&amp;gt;|

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch 规则


**ItemHasRegularExpressionMatch** 规则对于基于受支持属性的特定值控制外接程序的激活很有用。**ItemHasRegularExpressionMatch** 规则包括以下属性。



|**属性名**|**说明**|
|:-----|:-----|
|**RegExName**|指定正则表达式的名称，以便能够在外接程序的代码中引用该表达式。|
|**RegExValue**|指定将对其求值的正则表达式以确定是否应显示外接程序。|
|**PropertyName**|指定将为其对正则表达式求值的属性的名称。允许的值是  **BodyAsHTML**、 **BodyAsPlaintext**、 **SenderSMTPAddress** 和 **Subject**。 如果指定  **BodyAsHTML**，则 Outlook 仅在项目正文为 HTML 时应用正则表达式，否则，Outlook 将不会为该正则表达式返回任何匹配项。由于约会始终以 RTF 格式保存，因此，指定  **BodyAsHTML** 的正则表达式不与约会项目的正文中的任何字符串匹配。如果指定  **BodyAsPlaintext**，则 Outlook 始终对项目正文应用正则表达式。|
|**IgnoreCase**|指定在匹配由 **RegExName** 指定的正则表达式时是否忽略大小写。|

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>在规则中使用正则表达式的最佳实践

在使用正则表达式时应特别注意以下几点：


- 如果对项目的正文指定  **ItemHasRegularExpressionMatch** 规则，则正则表达式将进一步筛选正文且不应尝试返回项目的整个正文。使用正则表达式（如 `.*`）尝试获取项目的整个正文并不会始终返回预期结果。
    
- 各个浏览器上返回的纯文本正文可能存在细微的差异。如果使用 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 规则（将 **BodyAsPlaintext** 作为 **PropertyName** 属性），请在您的加载项支持的所有浏览器上测试您的正则表达式。
    
    因为不同的浏览器获取所选项目的文本正文的方法不同，所以应确保你的正则表达式支持正文文本部分所返回的细微差异。例如，一些浏览器（例如 Internet Explorer 9）使用 DOM 的 **innerText** 属性，而其他浏览器（例如 Firefox）使用.**textContent()** 方法来获取项目的文本正文。同样，不同的浏览器所返回的换行符可能不同：在 Internet Explorer 上所返回的换行符为“\r\n”，而在 Firefox 和 Chrome 上所返回的换行符为“\n”。有关详细信息，请参阅 [W3C DOM 兼容性 - HTML](http://www.quirksmode.org/dom/w3c_html.mdl#t07)。
    
- Outlook 富客户端与 Outlook Web App 或 适用于设备的 OWA 上的项目的 HTML 正文之间存在细微差异。请仔细定义您的正则表达式。例如，请考虑在  **ItemHasRegularExpressionMatch** 规则中（将 **BodyAsHTML** 作为 **PropertyName** 属性值）使用的以下正则表达式：
    
```
      http.*\.contoso\.com
```


    A rule with this regular expression would match the string "http-equiv="Content-Type" which exists in the HTML body of an item in an Outlook rich client, as part of the following  **META** tag:
    

```HTML
      <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
```


由于这些主机中的 HTML 正文不包含 **META** 标记，因此，相同的规则不会在 Outlook Web App 和适用于设备的 OWA 中返回此匹配项。这会影响是否针对不同的 Outlook 客户端适当地激活外接程序。在此示例中，请改用以下正则表达式：
    

```
      http://.*\.contoso\.com/
```

- 根据主机应用程序、设备类型或将对其应用正则表达式的属性，在设计正则表达式作为激活规则时，您需要了解针对每个主机的其他最佳实践和限制。有关详细信息，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)。
    

### <a name="examples"></a>示例

以下  **ItemHasRegularExpressionMatch** 规则将在发件人的 SMTP 电子邮件地址与"@contoso"匹配（不管是大写还是小写字符）时激活外接程序。


```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]" 
    PropertyName="SenderSMTPAddress"
/>
```

以下是使用  **IgnoreCase** 属性指定同一正则表达式的另一种方式。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@contoso" 
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

以下  **ItemHasRegularExpressionMatch** 规则将在股票代号包含在当前项目的正文中时激活外接程序。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    PropertyName="BodyAsPlaintext" 
    RegExName="TickerSymbols" 
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```


## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity 规则



  **ItemHasKnownEntity** 规则根据所选项目的主题和正文中是否存在实体来激活外接程序。[KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) 类型定义了受支持的实体。对 **ItemHasKnownEntity** 规则应用正则表达式将有助于基于实体的一部分值（例如，一组特定的 URL 或包含特定区号的电话号码）进行激活。


 >
  **注释**  Outlook 只能提取用英语编写的实体字符串，无论清单中指定的默认区域设置如何。 仅邮件而非约会支持  **MeetingSuggestion** 实体类型。您无法从"已发送邮件"文件夹的项目中提取实体，也不能使用 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 规则激活针对"已发送邮件"文件夹中的项目的外接程序。

**ItemHasKnownEntity** 规则支持下表中的属性。请注意，在 **ItemHasKnownEntity** 规则中指定正则表达式是可选的，如果你选择将正则表达式用作实体筛选器，则必须同时指定 **RegExFilter** 和 **FilterName** 属性。



|**属性名**|**说明**|
|:-----|:-----|
|**EntityType**|指定必须为其计算结果为  **true** 的规则找到的实体的类型。使用多个规则可指定多个类型的实体。|
|**RegExFilter**|指定用于进一步筛选由  **EntityType** 指定的实体的实例的正则表达式。|
|**FilterName**|指定由 **RegExFilter** 指定的正则表达式的名称，以便稍后可通过代码引用它。|
|**IgnoreCase**|指定在匹配由 **RegExFilter** 指定的正则表达式时是否忽略大小写。|

### <a name="examples"></a>示例

以下  **ItemHasKnownEntity** 规则在当前项目的主题或正文中存在 URL 且该 URL 包含字符串"youtube"时将激活外接程序，而不考虑字符串的大小写。


```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```


## <a name="using-regular-expression-results-in-code"></a>在代码中使用正则表达式结果


您可以通过在当前项目上使用以下方法来获取正则表达式的匹配项：


- [getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md) 为外接程序的 **ItemHasRegularExpressionMatch** 和 **ItemHasKnownEntity** 规则中指定的所有正则表达式返回当前项目中的匹配项。
    
- [getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md) 为外接程序的 **ItemHasRegularExpressionMatch** 规则中指定的已标识正则表达式返回当前项目中的匹配项。
    
- [getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) 返回包含外接程序的 **ItemHasKnownEntity** 规则中指定的已标识正则表达式的匹配项的实体的完整实例。
    
在对正则表达式求值时，匹配项将返回到数组对象中的外接程序。对于  **getRegExMatches**，该对象具有正则表达式的名称的标识符。 


 >**注释**  Outlook 富客户端不会以任何特定顺序返回数组中的匹配项。此外，您不应假定 Outlook 富客户端将以同一顺序返回该数组中 Outlook Web App 或 适用于设备的 OWA 的匹配项，即使您对每个此类客户端上的同一邮箱中的相同项目运行同一个外接程序也是如此。有关在 Outlook 富客户端和 Outlook Web App 或 适用于设备的 OWA 上处理正则表达式的方式的其他差异，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)。


### <a name="examples"></a>示例

以下是包含带名为  `videoURL` 的正则表达式的 **ItemHasRegularExpressionMatch** 规则的规则集的示例。


```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="VideoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="Body"/>
</Rule>
```

以下示例使用当前项目的  **getRegExMatches** 将变量 `videos` 设置为上一个 **ItemHasRegularExpressionMatch** 规则的结果。




```
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

多个匹配项将作为数组元素存储在该对象中。以下代码示例说明如何对名为  `reg1` 的正则表达式循环访问匹配项以生成将显示为 HTML 的字符串。




```js
function initDialer() 
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = _Item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }
    myCell.innerHTML = myString;
}

```

以下是  **ItemHasKnownEntity** 规则的示例，它指定 **MeetingSuggestion** 实体和名为 `CampSuggestion` 的正则表达式。Outlook 会在检测到当前所选项目包含会议建议且主题或正文包含"WonderCamp"一词时激活外接程序。




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

以下代码示例使用当前项目的  **getFilteredEntitiesByName(name)** 设置变量 `suggestions` 以获取针对上一个 **ItemHasKnownEntity** 规则的一系列已检测到的会议建议。




```
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName(CampSuggestion);
```


## <a name="additional-resources"></a>其他资源



- [创建适用于阅读窗体的 Outlook 外接程序](../outlook/read-scenario.md)
    
- [Outlook 外接程序的激活规则](../outlook/manifests/activation-rules.md)
    
- [Outlook 外接程序的激活和 JavaScript API 限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [将 Outlook 项目中的字符串作为已知实体进行匹配](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- 
  [.NET Framework 中的正则表达式的最佳做法](http://msdn.microsoft.com/en-us/library/618e5afb-3a97-440d-831a-70e4c526a51c%28Office.15%29.aspx)
    
