منظومة التحليل الشامل والتوجيه النهائي لنتائج الفصل الثالث والسنوي (الإصدار 03)
Term 03 & Annual Results Analysis, Experimental BAC & Final Orientation Suite (v3.0)
منظومة برمجية متكاملة مطورة بلغة VBA for Excel (Office 365) ومصممة خصيصاً لمستشاري التوجيه والإرشاد المدرسي والمهني وإدارات التعليم الثانوي بالجزائر. تقوم المنظومة بمعالجة ملفات الرقمنة الخام للفصل الثالث والمعدلات السنوية، وإصدار مصنفات التحليل الإحصائي، وبطاقات المتابعة الفردية، وتحليل نتائج البكالوريا التجريبية، وإدارة سلسلة التوجيه النهائي لتلاميذ السنة الأولى ثانوي (جذع مشترك علوم وتكنولوجيا وجذع مشترك آداب).
---
المزايا المعمارية في الإصدار الثالث (v3.0)
استقلالية تامة للوحدات (Zero-Coupling Architecture):
تقرأ الوحدات التنفيذية (`modAverage03`، `modSplitTermResults`، `modTransfer03`، ومجموعة `modAnalysisBuilder*`) مباشرة من المجلد المصدري `\Source03\` في الذاكرة العشوائية (`In-Memory Collection`) دون التوقف على وجود الملف الوسيط `Term03Results.xlsx`.
تعطل أو عدم الحاجة لأي مخرج لا يعطل بقية الأزرار أو الوحدات المستقلة.
لوحة تحكم موحدة (`ControlPanel`):
دمج كامل للوحة متابعة ملفات الشعب الـ 12 (`Term03Panel` سابقاً) مع لوحة أزرار العمليات الـ 8 في ورقة واحدة متناسقة (`ControlPanel`).
تحديث حي لحالة بناء كل ملف وزمن المعالجة دون الحاجة لملفات سجل نصية خارجية (`BuildLog` ملغى).
تشفير نصي نقي (Pure ASCII / Hex Code-Points):
جميع النصوص العربية والتسميات الرسمية مخزنة كرموز ست عشرية (`Hex Code Points`) عبر الدوال `ArT` و`Txt` و`ArHex`، مما يحمي الكود البرمجي من تلف الترميز (`Mojibake`) عند نقله بين أنظمة تشغيل ذات إعدادات لغة إقليمية مختلفة.
تحصين شامل للمعادلات والملفات:
تغليف كافة المعادلات المولدة برمجياً (`COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `SUMPRODUCT`, `MODE.SNGL`, `AGGREGATE`, `INDEX`) بدوال `IFERROR`.
نظام عدم الاستبدال التلقائي (`_01`, `_02`) لملفات التحليل، ونظام النسخ الاحتياطي التلقائي (`_Backup`) لملفات الفرز والمعدلات.
---
هيكل المجلدات والملفات
يجب وضع مصنف الأداة الرئيسي في مجلد محلي (غير سحابي مباشر `http`)، وبجانبه المجلدات والملفات التالية:
```text
📁 جذر المشروع (بجانب مصنف الأداة الرئيسي)
 ├── 📂 Source03\                        # يضم ملفات الرقمنة الخام لجميع الأفواج (قراءة فقط أو أرشفة تلقائية في _Original)
 ├── 📂 ExamenSource03\                  # يضم ملفات نقاط البكالوريا التجريبية لكل فوج ومادة (.xlsx)
 ├── 📂 تحليل نتائج الفصل الثالث\         # يُنشأ تلقائياً ويضم مصنفات التحليل الـ 12 لكل المستويات والشعب
 ├── 📂 _Backup\                         # يُنشأ تلقائياً لحفظ النسخ السابقة من ملفات الفرز والمعدلات
 ├── 📄 Term03Results.xlsx               # المجمع العام لنتائج الفصول الثلاثة (يصدر عن زر Third Term Results)
 ├── 📄 AnnualAverages.xlsx              # مصنف المعدلات السنوية لجميع الأفواج (يصدر عن زر Annual Averages)
 ├── 📄 FirstTermOnly.xlsx               # نتائج الفصل الأول فقط (يصدر عن زر Split Term Results)
 ├── 📄 SecondTermOnly.xlsx              # نتائج الفصل الثاني فقط (يصدر عن زر Split Term Results)
 ├── 📄 ThirdTermOnly.xlsx               # نتائج الفصل الثالث فقط (يصدر عن زر Split Term Results)
 ├── 📄 تحليل نتائج البكالوريا التجريبية.xlsx # يصدر عن زر Examen Analysis
 ├── 📄 FinalStreamPreferences.xlsx      # مصنف مصدر رغبات التلاميذ (ورقتي Tecno و Litr)
 └── 📄 FinalGuidance.xlsx               # مصنف التوجيه النهائي (Data، Data01، حوصلة عامة، الحوصلة)
```
---
قائمة الوحدات النمطية المعتمدة (12 Modules)
رقم	اسم الوحدة النمطية (`VB_Name`)	الدور الوظيفي في المنظومة
1	`modBuildPanel03`	بناء وإدارة لوحة التحكم الموحدة `ControlPanel` وجدول حالة الملفات وسلسلة الأمان
2	`modCollector`	التحويل الصامت للملفات القديمة وجمع `\Source03\` في `Term03Results.xlsx`
3	`modAnalysisBuilderCommon03`	الثوابت العامة، الخطوط، الأنماط، واكتشاف هيكل البيانات (`DiscoverData`)
4	`modAnalysisBuilderData03`	قراءة `\Source03\` وبناء ورقتي `Data` و`Data01` لمصنفات الشعب الـ 12
5	`modAnalysisBuilderAnalysis03`	بناء ورقتي `Analysis03` و`AnnualAnalysis` مع الاحتواء التلقائي `AutoFit`
6	`modAnalysisBuilderCartTec03`	بناء البطاقة الفنية الفردية للتلميذ `Cart_Tec` المصممة للعرض (`DataShow`)
7	`modAnalysisBuilderBuildAll03`	الماكرو الرئيسي `Build_All` و`Build_One` لتوليد مصنفات الشعب الـ 12
8	`modAverage03`	قراءة `\Source03\` مباشرة وحساب المعدلات السنوية في `AnnualAverages.xlsx`
9	`modSplitTermResults`	قراءة `\Source03\` وفرز الفصول الثلاثة دفعة واحدة في 3 مصنفات مستقلة
10	`modExamenAnalysis03`	قراءة `\ExamenSource03\` وإصدار مصنف تحليل البكالوريا التجريبية
11	`modTransfer03`	قراءة `\Source03\` مباشرة وتحويل معطيات ومعدلات وملامح التوجيه إلى `FinalGuidance.xlsx`
12	`modStreamPreferences03` + `modOrientation03`	استيراد الرغبات (`FinalStreamPreferences.xlsx`)، تنفيذ خوارزمية التوجيه، وبناء التقارير الحية
---
متطلبات التشغيل
نظام التشغيل: Windows 10 أو Windows 11.
إصدار الأوفيس: Microsoft Excel 365 أو Excel 2021 فما فوق (لدعم دوال `MAXIFS`، `MINIFS`، `MODE.SNGL`، و`Formula2` للمصفوفات الديناميكية).
الخطوط المفضلة: `Dubai`، `Dubai Medium`، `Aptos Slab`، `Cairo`، `Bahnschrift SemiBold` (مع وجود نظام استبدال تلقائي ذكي بالخطوط القياسية `Arial` و`Calibri` في حال عدم توفرها).
