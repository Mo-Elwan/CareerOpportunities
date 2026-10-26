<#
====================================================================================================
  وزارة البيئة والمياه والزراعة - MEWA
  سكربت إنشاء قائمة شيربوينت: نموذج تسجيل الاهتمام بالعمل (قاعدة الكفاءات المستقبلية)
  SharePoint List Provisioning Script: CareerInterests
====================================================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "رابط موقع شيربوينت للوزارة")]
    [string]$SiteUrl = "https://www.mewa.gov.sa/ar/Ministry/AboutMinistry/Employment/CareerOpportunities",

    [Parameter(Mandatory = $false)]
    [string]$ListName = "CareerInterests",

    [Parameter(Mandatory = $false)]
    [string]$ListTitleArabic = "سجل الاهتمام بالعمل - قاعدة الكفاءات"
)

Write-Host "==========================================================" -ForegroundColor Green
Write-Host " بدء تهيئة قائمة شيربوينت: $ListTitleArabic" -ForegroundColor Yellow
Write-Host " الموقع: $SiteUrl" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Green

# 1. التحقق من تثبيت وحدة PnP PowerShell
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Warning "وحدة PnP.PowerShell غير مثبتة. يمكنك تثبيتها عبر الأمر: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
    Write-Host "محاولة التثبيت الآن..." -ForegroundColor Yellow
    try {
        Install-Module -Name "PnP.PowerShell" -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "تم تثبيت PnP.PowerShell بنجاح." -ForegroundColor Green
    }
    catch {
        Write-Error "تعذر تثبيت PnP.PowerShell تلقائياً. يرجى تشغيل السكربت كمسؤول أو مراجعة الدليل اليدوي المرفق."
        exit
    }
}

Import-Module PnP.PowerShell

# 2. الاتصال بموقع شيربوينت
try {
    Write-Host "جارٍ الاتصال بموقع شيربوينت..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive
    Write-Host "تم الاتصال بنجاح!" -ForegroundColor Green
}
catch {
    Write-Error "فشل الاتصال بموقع شيربوينت: $_"
    exit
}

# 3. التحقق من وجود القائمة أو إنشائها
$existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

if ($null -eq $existingList) {
    Write-Host "إنشاء القائمة الجديدة: $ListName ..." -ForegroundColor Cyan
    $list = New-PnPList -Title $ListName -Template GenericList -EnableVersioning
    Set-PnPList -Identity $ListName -Title $ListTitleArabic -Description "قائمة تسجيل اهتمام الكفاءات والكوادر المهنية للانضمام لوزارة البيئة والمياه والزراعة"
    Write-Host "تم إنشاء القائمة بنجاح." -ForegroundColor Green
} else {
    Write-Host "القائمة موجودة مسبقاً: $ListName. سيتم إضافة أو تحديث الحقول." -ForegroundColor Yellow
    $list = $existingList
}

# 4. إعادة تسمية حقل Title الافتراضي إلى "الاسم الثلاثي"
Set-PnPField -List $ListName -Identity "Title" -Values @{ Title = "الاسم الثلاثي"; Required = $true }
Write-Host "تم ضبط حقل Title كـ 'الاسم الثلاثي'." -ForegroundColor Green

# 5. إضافة الحقول المطلوبة وفق ملاحظات إدارة الاستقطاب
$fields = @(
    # البيانات الأساسية والتواصل
    @{ Name = "NationalID"; DisplayName = "رقم الهوية"; Type = "Text"; Required = $true },
    @{ Name = "Nationality"; DisplayName = "الجنسية"; Type = "Choice"; Choices = @("سعودي", "غير سعودي"); Required = $true },
    @{ Name = "Email"; DisplayName = "البريد الإلكتروني"; Type = "Text"; Required = $true },
    @{ Name = "MobileNumber"; DisplayName = "رقم الجوال"; Type = "Text"; Required = $true },
    @{ Name = "Region"; DisplayName = "العنوان: المنطقة"; Type = "Choice"; Choices = @("منطقة الرياض", "منطقة مكة المكرمة", "المنطقة الشرقية", "منطقة المدينة المنورة", "منطقة القصيم", "منطقة عسير", "منطقة تبوك", "منطقة حائل", "منطقة الحدود الشمالية", "منطقة جازان", "منطقة نجران", "منطقة الباحة", "منطقة الجوف"); Required = $true },
    @{ Name = "City"; DisplayName = "العنوان: المدينة"; Type = "Text"; Required = $true },

    # المؤهلات العلمية
    @{ Name = "EducationDegree"; DisplayName = "الدرجة العلمية"; Type = "Choice"; Choices = @("ثانوي فما دون", "دبلوم", "بكالوريوس", "ماجستير", "دكتوراه"); Required = $true },
    @{ Name = "Major"; DisplayName = "التخصص"; Type = "Text"; Required = $true },
    @{ Name = "University"; DisplayName = "اسم الجامعة"; Type = "Text"; Required = $true },
    @{ Name = "GraduationYear"; DisplayName = "سنة التخرج"; Type = "Text"; Required = $true },
    @{ Name = "GPAScale"; DisplayName = "مقياس المعدل"; Type = "Choice"; Choices = @("من 4", "من 5", "من 100"); Required = $true },
    @{ Name = "GPA"; DisplayName = "المعدل التراكمي"; Type = "Text"; Required = $true },

    # الخبرات والبيانات المهنية
    @{ Name = "ExperienceYears"; DisplayName = "سنوات الخبرة"; Type = "Choice"; Choices = @("دون خبرة", "أقل من سنة", "1 - 2 سنة", "3 - 5 سنوات", "6 - 10 سنوات", "أكثر من 10 سنوات"); Required = $true },
    @{ Name = "EmploymentStatus"; DisplayName = "الحالة الوظيفية الحالية"; Type = "Choice"; Choices = @("على رأس العمل", "باحث عن عمل"); Required = $true },
    @{ Name = "CurrentJobTitle"; DisplayName = "المسمى الوظيفي الحالي"; Type = "Text"; Required = $false },
    @{ Name = "CurrentEmployer"; DisplayName = "جهة العمل الحالية"; Type = "Text"; Required = $false },

    # اللغة والمصادقة
    @{ Name = "EnglishLevel"; DisplayName = "مستوى اللغة الإنجليزية"; Type = "Choice"; Choices = @("مبتدئ", "متوسط", "متقدم"); Required = $true },
    @{ Name = "ConsentDeclaration"; DisplayName = "الإقرار والموافقة"; Type = "Boolean"; Required = $true }
)

foreach ($f in $fields) {
    $existing = Get-PnPField -List $ListName -Identity $f.Name -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        Write-Host "إضافة الحقل: $($f.DisplayName) ($($f.Name)) ..." -ForegroundColor Gray
        if ($f.Type -eq "Choice") {
            Add-PnPField -List $ListName -InternalName $f.Name -DisplayName $f.DisplayName -Type Choice -Choices $f.Choices -AddToDefaultView
        }
        elseif ($f.Type -eq "Boolean") {
            Add-PnPField -List $ListName -InternalName $f.Name -DisplayName $f.DisplayName -Type Boolean -AddToDefaultView
        }
        else {
            Add-PnPField -List $ListName -InternalName $f.Name -DisplayName $f.DisplayName -Type Text -AddToDefaultView
        }
        Set-PnPField -List $ListName -Identity $f.Name -Values @{ Required = $f.Required }
    } else {
        Write-Host "الحقل موجود: $($f.DisplayName)" -ForegroundColor DarkGray
    }
}

# 6. تفعيل المرفقات في القائمة (لحفظ السيرة الذاتية والشهادات)
Set-PnPList -Identity $ListName -EnableAttachments $true
Write-Host "تم التأكد من تفعيل المرفقات (Attachments) في القائمة لحفظ السير الذاتية." -ForegroundColor Green

Write-Host "==========================================================" -ForegroundColor Green
Write-Host " اكتملت تهيئة قائمة شيربوينت بنجاح!" -ForegroundColor Green
Write-Host " اسم القائمة البرمجي: $ListName" -ForegroundColor Cyan
Write-Host " يمكنك الآن استخدام ملف sharepoint-webpart.html في الصفحة." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Green
