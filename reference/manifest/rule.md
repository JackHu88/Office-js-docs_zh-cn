
# <a name="rule-element"></a>Rule 元素
指定应针对此邮件外接程序计算的激活规则。

 **外接程序类型：**邮件


## <a name="syntax:"></a>语法：

 **ItemIs Rule** - 定义一个在选定项为指定类型时计算结果为 true 的规则。


```XML
<Rule xsi:type="ItemIs" 
   ItemType= ["Appointment" | "Message"]
   FormType=["Read" | "Edit" | "ReadOrEdit"] 
   ItemClass = "string " 
   IncludeSubClasses=["true" | "false"] />
```

 **ItemHasAttachment Rule** - 定义一个当项目包含附件时计算结果为 True 的规则。




```XML
<Rule xsi:type="ItemHasAttachment"  />
```

 **ItemHasKnownEntity** - 定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。




```XML
<Rule xsi:type="ItemHasKnownEntity" 
  EntityType=["MeetingSuggestion" | "TaskSuggestion" |"Address" | "Url" | "PhoneNumber" | "EmailAddress" | "Contact" ]
  RegExFilter="string "
  FilterName="string "
  IgnoreCase=["true | false"]/>
```

 **ItemHasRegularExpressionMatch Rule** - 定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="string " 
    RegExValue="string " 
    PropertyName=["Subject" | "BodyAsPlaintext" | "BodyAsHtml" | "SenderSTMPAddress"]
    IgnoreCase=["true" | "false"]
/>
```

 **RuleCollection Rule** - 定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。




```XML
<Rule xsi:type="RuleCollection" Mode=["And" | "Or"]>
   ...
</Rule>
```


## <a name="contained-in:"></a>包含在：

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## <a name="attributes:"></a>属性：

 **ItemIs Rule 属性**



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|ItemType|ItemType（字符串）|必需|指定要匹配的项目类型。可以是下列类型之一：

|**ItemType**|**Corresponding ItemClass**|
|:-----|:-----|
|Appointment|IPM.Appointment|
|Message(1)|包括电子邮件、会议请求、响应和取消。|
|
|FormType|ItemFormType（字符串）|必需|指定应用应出现在项目的读取还是编辑表单中。可以是下列类型之一。|

|**FormType**|**说明**|
|:-----|:-----|
|Read|指定仅在（指定了 **ItemType** 的）阅读窗体中激活邮件外接程序。|
|Edit|指定仅在（指定了 **ItemType** 的）撰写窗体中激活邮件外接程序。|
|ReadOrEdit|指定在（指定了 **ItemType** 的）阅读和撰写窗体中激活邮件外接程序。|
|ItemClass|字符串|可选|指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件外接程序](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx)。|
|IncludeSubClasses|布尔|可选|指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 false。|


(1) 下面是相应的邮件类别：IPM.NoteIPM.Schedule.Meeting.RequestIPM.Schedule.Meeting.NegIPM.Schedule.Meeting.PosIPM.Schedule.Meeting.TentIPM.Schedule.Meeting.Canceled。

 **ItemHasAttachment Rule 属性**

无。

 **ItemHasKnownEntity Rule 属性**



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|EntityType|KnownEntityType（字符串）|必需|指定若想规则计算结果为 true 而必须存在的实体类型。可以是下列类型之一：

|**KnownEntityType**|**Descripition**|
|:-----|:-----|
|MeetingSuggestion|由模式识别功能识别的引用事件或会议的文本。|
|TaskSuggestion| 由模式识别功能识别的包含可操作短语的文本。|
|Address|由模式识别功能识别的引用美国邮政地址的文本。|
|Url|由模式识别功能识别的包含文件名或 Web 地址 URL 的文本。|
|PhoneNumber| 由模式识别功能识别为北美洲电话号码的一系列数字。|
|EmailAddress|由模式识别功能识别的包含 SMTP 格式的电子邮件地址的文本。|
|Contact|由模式识别功能识别的包含联系人信息的文本。|
|RegExFilter|字符串|可选|指定一个针对此实体运行以进行激活的正则表达式。|
|FilterName|字符串|可选|指定正则表达式筛选器的名称，以便随后能够在您的外接程序代码中引用该名称。|
|IgnoreCase|布尔|可选|指定在运行由 **RegExFilter** 属性指定的正则表达式时忽略大小写。|
 **ItemHasRegularExpressionMatch Rule 属性**



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|RegExName|字符串|必需|指定正则表达式的名称，以便您能够在外接程序的代码中引用该表达式。|
|RegExValue|字符串|必需|指定将对其求值的正则表达式以确定是否应显示邮件外接程序。 |
|PropertyName|PropertyName（字符串）|必需|指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：

|**PropertyName**|**说明**|
|:-----|:-----|
|Subject|根据项目主题计算正则表达式。|
|BodyAsPlaintext|根据纯文本形式的项目正文计算正则表达式。|
|BodyAsHtml|根据项目正文（如果正文采用 HTML 格式）计算正则表达式。|
|SenderSTMPAddress|根据项目发件人的 SMTP 地址计算正则表达式。|
|IgnoreCase|布尔|可选|指定在执行正则表达式时忽略大小写。|
 **RuleCollection Rule 属性**



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|Mode|字符串|必需|指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：“And”或“Or”。|

## <a name="additional-resources"></a>其他资源



- 
  [在 Outlook 中为特定邮件类别激活邮件外接程序](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx) 和 [Outlook 外接程序的激活规则](../../docs/outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)
    
- [将 Outlook 项目中的字符串作为已知实体进行匹配](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [使用正则表达式激活规则显示 Outlook 外接程序](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
