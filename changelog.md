# سجل التغييرات والتحديثات (Changelog)

جميع التغييرات الجوهرية والمعمارية التي طرأت على منظومة تحليل معطيات الفصل الثالث والتوجيه النهائي موثقة في هذا الملف.

---

## [الإصدار 3.0.0] - 2026-11-05 (الهيكلة الموحدة وفك الارتباط الكامل)

### تغييرات معمارية كبرى (Breaking & Architectural Changes)
* **فك الارتباط بالملف الوسيط `Term03Results.xlsx`:**
  * تم تضمين محرك التجميع المباشر من المجلد المصدري `\Source03\` داخل كل من `modAverage03` و`modTransfer03` و`modSplitTermResults`.
  * أصبحت جميع الوحدات تعمل باستقلالية تامة؛ فشل أي وحدة أو عدم تشغيل الزر الأول (`Third Term Results`) لا يوقف عمل بقية الوحدات.
* **دمج وحدات فرز الفصول في `modSplitTermResults`:**
  * دمج الوحدات الثلاث (`FirstTermOnly1`، `SecondTermOnly1`، `ThirdTermOnly`) في وحدة نمطية واحدة تحت مسمى `modSplitTermResults`.
  * استبدال الأزرار الثلاثة في لوحة التحكم بزر واحد **`Split Term Results`** يستخرج الفصول الثلاثة دفعة واحدة من `\Source03\`.
* **دمج لوحتي التحكم في `ControlPanel` (`modBuildPanel03`):**
  * نقل كافة بيانات ومعلومات وجدول حالة الملفات من ورقة `Term03Panel` إلى لوحة التحكم الرئيسية `ControlPanel` مع الحذف التلقائي لورقة `Term03Panel` والوحدة `modAnalysisBuilderPanel03`.
  * إدراج الزر **`Term 03 Analysis`** في شبكة الأزرار مكان الزر الملغى `Annual Analysis`.
  * حذف زر "فتح مجلد الإصدار" والأسطر التي كانت تضمه، ورفع جدول **"حالة الملفات"** وعنوانه من السطر 14 إلى **السطر 10**.

### إضافات وتحسينات (Added & Improved)
* **تسمية أوراق التحليل والاحتواء التلقائي (`modAnalysisBuilderCommon03` & `modAnalysisBuilderAnalysis03`):**
  * تغيير اسم ورقة تحليل الفصل الثالث من `Analysis01` إلى **`Analysis03`**.
  * تطبيق الاحتواء التلقائي لعرض الأعمدة (`AutoFit`) على النطاق من العمود `D` إلى آخر عمود بيانات في ورقتي `Analysis03` و`AnnualAnalysis` بعد إتمام الحساب التلقائي للصيغ، مع الحفاظ على حد أدنى للعرض (`6.3`) لمنع تضاغط العناوين العمودية.
* **تحصين المعادلات (`IFERROR` Wrapping):**
  * تغليف كافة صيغ `COUNTIFS` و`SUMPRODUCT` وروابط الخلايا المباشرة في جداول `Analysis03` و`AnnualAnalysis` بدالة `IFERROR` لضمان خلو الجداول من أي أخطاء إكسيل مرئية.
* **ضبط وتوحيد أسماء الوحدات النمطية (`VB_Name` < 31 حرفاً):**
  * تقصير أسماء وحدات البناء الطويلة (`modAnalysisWorkbookBuilder*` التي كانت تتجاوز 31 حرفاً) إلى `modAnalysisBuilderCommon03`، `modAnalysisBuilderData03`، `modAnalysisBuilderAnalysis03`، `modAnalysisBuilderCartTec03`، `modAnalysisBuilderBuildAll03`.
* **إزالة التكرار البرمجي (`Code Deduplication`):**
  * توحيد الدوال المساعدة في `modAnalysisBuilderData03` لتستخدم مباشرة الدوال العامة في `modAnalysisBuilderCommon03` (`ArT`، `NormKey`، `FileExistsP`، `FolderExistsP`، `IsOpenWb`).

### إصلاحات (Fixed)
* إصلاح مسار استدعاء زر البكالوريا التجريبية `Examen Analysis` في `modBuildPanel03` ليشير إلى `modExamenAnalysis03.Run_ExamenAnalysis03`.
* إصلاح مسارات `OnAction` التي كانت تشير إلى أسماء وحدات قديمة بها أخطاء إملائية (`modBiuldAnalysisfiles...`).
* تنظيف كافة الوحدات من علامات التنصيص المزدوجة الزائدة وأسطر الترويسة غير المعلّقة التي كانت تمنع الترجمة الفورية في محرر VBA.

### محذوفات (Removed)
* إلغاء توليد ملف السجل النصي `BuildLog_yyyymmdd_hhnnss.txt` وحذف دالة `WriteLog` من `modAnalysisBuilderBuildAll03`.
* حذف الزر `Annual Analysis` والوحدة المرتبطة به `modResultFlow03`.
* حذف الزر `Clear Tables` والوحدة المرتبطة به `modClearDataOnly03`.
* حذف الوحدات المنفصلة `FirstTermOnly1` و`SecondTermOnly1` و`ThirdTermOnly` بعد دمجها في `modSplitTermResults`.
* حذف الوحدة `modAnalysisWorkbookBuilderPanel` (`modAnalysisBuilderPanel03`) بعد دمجها الكامل في `modBuildPanel03`.