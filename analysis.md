# ניתוח: Excel Data Interface – Urban Renewal AI

## מטרה
פיתוח מנגנון Python לקריאה וכתיבה של קובצי Excel המשמשים לאיסוף, ניתוח ואימות מידע סטטוטורי
במתחמי התחדשות עירונית בישראל. הממשק פועל כשכבת נתונים בין סוכן AI לבין מאגר מובנה.

---

## מבנה הקבצים בתיקייה

```
2026.11.02/
├── אפיון דרישות.docx                      ← מסמך דרישות מערכת מלא
├── ארכיטקטורה טכנולוגית.docx              ← ארכיטקטורה רב-סוכנים על IBM watsonx
├── מיזם-טכנולוגי-UrbanAI.pdf              ← מסמך מיזם
├── נתונים_תכנוניים_הנדסיים_סוכן_1.xlsx   ← פירוק משימות לסוכן 1
├── סוכן -1-איסוף-מידע/
│   └── איסוף_נתונים_מתחם_סוכן_אוטומטי.xlsx
└── Excell/
    ├── Readme.xlsx                         ← כללי עבודה מחייבים
    ├── INSTRUCTIONS.md.docx                ← הוראות מפורטות לממשק
    ├── DATA_MAPPING.xlsx                   ← מילון שדות + Field_ID לכל שדה
    ├── גיליון SITE_DATA.xlsx               ← נתוני מתחם (tbl_site_data)
    ├── גיליון PARCELS.xlsx                 ← חלקות (tbl_parcels)
    ├── גיליון PLANNING_DATA.xlsx           ← תכניות (tbl_plans)
    ├── גיליון EXISTING_BUILDINGS.xlsx      ← מבנים (tbl_buildings)
    └── גיליון SOURCES_LOG.xlsx             ← מקורות ואימות (tbl_sources)
```

---

## מבנה ה-Excel (8 גיליונות)

### קשרים בין גיליונות

```
SOURCES_LOG (tbl_sources)
    Source_ID ──────────────────────────────────────────────────────┐
                                                                    │
SITE_DATA (tbl_site_data)                                          │
    Field_ID | Value                                                │
                                                                    │
PARCELS (tbl_parcels)                                              │
    Parcel_ID ← System │ Block_Number │ Parcel_Number │ Source_ID ←┤
         │                                                          │
         ├──→ EXISTING_BUILDINGS (tbl_buildings)                   │
         │         Building_ID │ Parcel_ID │ Floors │ Units │ Source_ID ←┤
         │                                                          │
         └──→ CONSTRAINTS (tbl_constraints)                        │
                   Constraint_ID │ Parcel_ID │ Type │ Source_ID ←──┘

PLANNING_DATA (tbl_plans)
    Plan_ID │ Plan_Number │ Plan_Status │ Land_Use │ Source_ID ←────┘

DATA_MAPPING (מילון)
    Field_ID → Sheet_Name → Table_Name → Column_Name → Type → Write_Permission
    README — הוראות ידניות
```

---

## מילון שדות — DATA_MAPPING

| Field_ID | שם שדה | גיליון | טבלה | עמודה | סוג | Write? |
|----------|--------|--------|------|-------|-----|--------|
| SITE_001 | שם עיר | SITE_DATA | tbl_site_data | Value | text | כן |
| SITE_002 | שכונה | SITE_DATA | tbl_site_data | Value | text | כן |
| SITE_003 | כתובת/ציר | SITE_DATA | tbl_site_data | Value | text | כן |
| SITE_005 | שטח מתחם | SITE_DATA | tbl_site_data | Value | number | כן |
| PARCEL_001 | מזהה חלקה | PARCELS | tbl_parcels | Parcel_ID | text | לא (System) |
| PARCEL_002 | מספר גוש | PARCELS | tbl_parcels | Block_Number | text | כן |
| PARCEL_003 | מספר חלקה | PARCELS | tbl_parcels | Parcel_Number | text | כן |
| PARCEL_004 | שטח חלקה | PARCELS | tbl_parcels | Area_sqm | number | כן |
| PLAN_001 | מספר תכנית | PLANNING_DATA | tbl_plans | Plan_Number | text | כן |
| PLAN_002 | סטטוס תכנית | PLANNING_DATA | tbl_plans | Plan_Status | text | כן |
| PLAN_003 | ייעוד קרקע | PLANNING_DATA | tbl_plans | Land_Use | text | כן |
| BUILD_001 | קומות | EXISTING_BUILDINGS | tbl_buildings | Floors_Above_Ground | integer | כן |
| BUILD_002 | יח"ד | EXISTING_BUILDINGS | tbl_buildings | Dwelling_Units | integer | כן |
| BUILD_003 | שנת בנייה | EXISTING_BUILDINGS | tbl_buildings | Year_Built | integer | כן |
| CONST_001 | סוג מגבלה | CONSTRAINTS | tbl_constraints | Constraint_Type | text | כן |
| CONST_002 | תיאור מגבלה | CONSTRAINTS | tbl_constraints | Description | text | כן |
| SRC_001 | שם מקור | SOURCES_LOG | tbl_sources | Source_Name | text | כן |

---

## מבנה הפתרון ב-Python

### 7 שלבי ביצוע ומרכיבי הקוד

```
שלב 1 — תכנון ארכיטקטורה   → הגדרת מחלקות ו-interfaces
שלב 2 — יצירת תבנית Excel  → יצירת קובץ Excel ריק מובנה לפי הגיליונות
שלב 3 — ExcelReader         → קריאת קובץ קיים, בניית מיפוי שדות
שלב 4 — DataMapper          → Field_ID → שדה ← נתון
שלב 5 — ExcelWriter         → כתיבה, הוספת שורות, עדכון רשומות
שלב 6 — DataValidator       → אימות טיפוסים, חיפוש סתירות, סטטוס
שלב 7 — בדיקות              → קריאה, כתיבה, עדכון, מניעת כפילויות
```

### ארכיטקטורת מחלקות

```
ExcelDataInterface (מתאם ראשי)
│
├── ExcelReader
│   ├── load_workbook(path)
│   ├── read_data_mapping()          ← תמיד ראשון
│   ├── read_sheet(sheet_name)
│   ├── identify_missing_fields()
│   └── get_existing_record(id)
│
├── DataMapper
│   ├── field_id_to_location(field_id)
│   ├── map_input_to_fields(data_dict)
│   └── validate_write_permission(field_id)
│
├── ExcelWriter
│   ├── write_field(field_id, value, source_id)
│   ├── append_row(sheet, table, row_data)
│   ├── update_row(sheet, table, id_field, id_value, updates)
│   └── save(output_path)
│
├── DataValidator
│   ├── validate_type(field_id, value)
│   ├── check_duplicates(sheet, id_field, id_value)
│   ├── detect_conflicts(field_id, new_value, existing_value)
│   └── generate_error_report()
│
└── SourcesLogger
    ├── log_source(source_name, source_type, date, reliability)
    ├── link_source_to_record(source_id, record_id)
    └── get_source_id(source_name)
```

---

## אופן הטיפול בכל תרחיש

| תרחיש | גישה |
|-------|------|
| **קריאת קובץ קיים** | `ExcelReader.load_workbook()` — טעינת כל הגיליונות, זיהוי טבלאות לפי שם בלבד (לא צבע/מיקום) |
| **קריאת DATA_MAPPING** | תמיד פעולה ראשונה לפני כל עיבוד — בניית מילון `{field_id: {sheet, table, column, type, write_permission}}` |
| **זיהוי שדות חסרים** | סריקת כל Field_ID בעלי `Required=yes` + `Write_Permission=yes` — רישום ב-missing_fields_report |
| **הוספת רשומה חדשה** | רק לגיליונות PARCELS / PLANNING_DATA / BUILDINGS / CONSTRAINTS — `append_row()` לאחר בדיקת כפילות |
| **עדכון רשומה קיימת** | חיפוש לפי ID ייחודי (Parcel_ID / Plan_ID / Building_ID / Constraint_ID) — עדכון שורה קיימת בלבד, לא מחיקה |
| **מניעת כפילויות** | לפני `append_row` — `check_duplicates()`. אם ID קיים → `update_row()` ולא `append_row()` |
| **תיעוד מקור** | כל כתיבה → `SourcesLogger.log_source()` + `link_source_to_record()`. כל ערך מקושר ל-Source_ID ב-SOURCES_LOG |
| **נתון סותר** | `DataValidator.detect_conflicts()` → סטטוס = `Conflicting Information`, רישום ב-`Conflict_Notes` |
| **שמירת קובץ** | קובץ מקורי → `_backup_YYYYMMDD.xlsx`. פלט → `_updated.xlsx` |
| **הרחבה עתידית** | מבנה מודולרי — כל גיליון הוא מחלקה נפרדת. הוספת עמודות/שדות → עדכון DATA_MAPPING בלבד |
| **דוח שגיאות** | `generate_error_report()` — שדות חסרים, שגיאות טיפוס, כפילויות, מקורות חסרים, כתיבה לא מורשית |

---

## סטטוסי נתונים (Data Status)

| ערך | משמעות |
|-----|--------|
| `Missing Information` | שדה חובה שלא מולא |
| `Collected` | נתון נאסף, לא אומת |
| `Requires Verification` | דורש אימות נוסף |
| `Verified` | אומת מול מקור רשמי |
| `Conflicting Information` | סתירה בין מקורות — דורש בדיקה |
| `Not Applicable` | לא רלוונטי למתחם הנוכחי |

---

## כללי עבודה מחייבים (מ-INSTRUCTIONS.md.docx ו-Readme.xlsx)

1. זיהוי שדה לפי: Field_ID → Sheet_Name → Table_Name → Column_Name (לא לפי צבע/מיקום ויזואלי)
2. לקרוא את DATA_MAPPING לפני כל עיבוד
3. לכתוב רק לשדות עם `Write_Permission = yes`
4. אין לדרוס נתונים קיימים ללא בדיקת ID
5. אין להמציא נתונים חסרים
6. אין לשנות מזהים קיימים (Parcel_ID, Plan_ID וכו')
7. לתעד מקור לכל נתון שנכתב
8. לזהות ולסמן נתונים סותרים
9. לשמור עותק מקורי לפני כל שינוי
10. להפריד בין `Collected` לבין `Verified`
11. להפיק דוח שגיאות + דוח שדות שלא מולאו

---

## מקורות נתונים לסוכן 1

| מקור | סוג | הערות |
|------|-----|--------|
| XPlan (מנהל התכנון) | API | תב"ע, זכויות בנייה |
| Govmap | GIS/API | שכבות מרחביות |
| מפ"י / Survey of Israel | API | גבולות חלקות, מדידות |
| GIS עירוני | API | נתוני עירייה |
| היתרי בנייה (RishunNet) | Web | תיקי בניין |
| לשכת רישום מקרקעין (טאבו) | בתשלום | נסחי טאבו, זכויות |

---

## הערות לפיתוח

- **openpyxl** — הספרייה המוגדרת. לא להשתמש בכלים חיצוניים (LibreOffice, pandas על קבצי pivot) למניעת השחתת נתונים.
- **Field_ID** הוא הצומת המרכזי — כל לוגיקה עוברת דרכו, לא דרך מיקום תא.
- **SOURCES_LOG** חייב להתעדכן בכל כתיבה — זה דרישת ה-audit trail של המערכת.
- גיליון **CONSTRAINTS** כרגע משתמש בשם טבלה `tbl_constraints` גם ב-SOURCES_LOG (יש לברר אם זו כוונה או שגיאה בתבנית).
